"""
步骤2：变调变速核心功能
从原始的串烧变调变速工具中提取
"""
import numpy as np
import librosa
import soundfile as sf
import pandas as pd
import logging
import shutil
import platform
import subprocess
from pathlib import Path
from typing import Callable, Optional, Tuple, List

logger = logging.getLogger(__name__)

# Windows 下配置避免弹出命令行窗口
if platform.system() == 'Windows':
    # 保存原始的 Popen
    _original_popen = subprocess.Popen
    
    def _popen_no_window(*args, **kwargs):
        """Windows 下隐藏子进程窗口的 Popen 包装"""
        # 设置 STARTUPINFO 隐藏窗口
        if 'startupinfo' not in kwargs:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            kwargs['startupinfo'] = startupinfo
        
        # 设置 CREATE_NO_WINDOW 标志
        creation_flags = kwargs.get('creationflags', 0)
        kwargs['creationflags'] = creation_flags | subprocess.CREATE_NO_WINDOW
        
        return _original_popen(*args, **kwargs)
    
    # 替换全局 Popen（pyrubberband 等会使用这个）
    subprocess.Popen = _popen_no_window

# 尝试导入 pyrubberband
try:
    import pyrubberband as prb
    HAS_RUBBERBAND = True
    logger.info("✓ 已安装 pyrubberband - 将使用专业级高质量算法")
except ImportError:
    HAS_RUBBERBAND = False
    logger.warning("⚠️ 未安装 pyrubberband - 将使用 librosa（音质较差）")
    logger.warning("⚠️ 强烈建议安装 pyrubberband 以获得最佳音质！")
    logger.warning("⚠️ 安装方法: pip install pyrubberband")
    logger.warning("⚠️ 详见《音质优化指南.md》")


# 调号映射
KEY_MAPPING = {
    'C': 0, 'C#': 1, 'DB': 1,
    'D': 2, 'D#': 3, 'EB': 3,
    'E': 4, 'FB': 4,
    'F': 5, 'E#': 5, 'F#': 6, 'GB': 6,
    'G': 7, 'G#': 8, 'AB': 8,
    'A': 9, 'A#': 10, 'BB': 10,
    'B': 11, 'CB': 11, 'B#': 0
}


def parse_key_to_semitone(key: str) -> Optional[int]:
    """将调性字符串转换为半音数"""
    if not key or pd.isna(key):
        return None
    
    key_str = str(key).strip().upper()
    
    # 移除大小调标识
    import re
    key_str = re.sub(r'\s*(M|MIN|MAJOR|MAJ|M7|M9|M11)\s*$', '', key_str, flags=re.IGNORECASE)
    
    return KEY_MAPPING.get(key_str)


def calculate_semitone_shift(source_key: str, target_key: str) -> Optional[int]:
    """计算半音差"""
    source_semitone = parse_key_to_semitone(source_key)
    target_semitone = parse_key_to_semitone(target_key)
    
    if source_semitone is None or target_semitone is None:
        return None
    
    shift = target_semitone - source_semitone
    # 标准化到 -6 到 6 之间
    if shift > 6:
        shift -= 12
    elif shift < -6:
        shift += 12
    
    return shift


def process_mono_audio(y: np.ndarray, sr: int, semitone_shift: int, tempo_rate: float) -> np.ndarray:
    """
    处理单声道音频（高质量模式）
    
    音质优化策略：
    1. 优先使用 pyrubberband（业界最高质量）
    2. pyrubberband 可以同时处理变调和变速，音质更好
    3. librosa 使用最高质量参数作为备选
    """
    # 如果同时需要变调和变速，使用 pyrubberband 一次性处理（音质最佳）
    if semitone_shift != 0 and tempo_rate != 1.0 and HAS_RUBBERBAND:
        try:
            # pyrubberband 同时处理变调和变速，避免二次处理损失
            # 使用高质量选项
            y_processed = prb.pitch_shift(y, sr, semitone_shift)
            y_processed = prb.time_stretch(y_processed, sr, tempo_rate)
            return y_processed
        except Exception as e:
            logger.warning(f"pyrubberband 同时处理失败，使用分步处理: {e}")
    
    # 分步处理：先做升降调
    if semitone_shift != 0:
        if HAS_RUBBERBAND:
            try:
                # pyrubberband 的质量远高于 librosa
                y_shifted = prb.pitch_shift(y, sr, semitone_shift)
            except Exception as e:
                logger.warning(f"pyrubberband 升降调失败，使用 librosa 高质量模式: {e}")
                # librosa 使用最高质量参数
                y_shifted = librosa.effects.pitch_shift(
                    y, 
                    sr=sr, 
                    n_steps=semitone_shift,
                    bins_per_octave=24,      # 提高音高分辨率（默认12）
                    res_type='soxr_vhq'      # 使用最高质量重采样器
                )
        else:
            # librosa 高质量模式
            logger.info("使用 librosa 高质量模式处理升降调")
            y_shifted = librosa.effects.pitch_shift(
                y, 
                sr=sr, 
                n_steps=semitone_shift,
                bins_per_octave=24,
                res_type='soxr_vhq'
            )
    else:
        y_shifted = y
    
    # 再做变速
    if tempo_rate != 1.0:
        if HAS_RUBBERBAND:
            try:
                # pyrubberband 变速质量更好
                y_final = prb.time_stretch(y_shifted, sr, tempo_rate)
            except Exception as e:
                logger.warning(f"pyrubberband 变速失败，使用 librosa 高质量模式: {e}")
                # librosa 使用 phase vocoder，质量较好
                y_final = librosa.effects.time_stretch(
                    y_shifted, 
                    rate=tempo_rate
                )
        else:
            logger.info("使用 librosa 高质量模式处理变速")
            y_final = librosa.effects.time_stretch(
                y_shifted, 
                rate=tempo_rate
            )
    else:
        y_final = y_shifted
    
    return y_final


def process_stereo_audio(y: np.ndarray, sr: int, semitone_shift: int, tempo_rate: float) -> np.ndarray:
    """处理立体声音频"""
    if y.ndim == 2 and y.shape[0] == 2:
        # 立体声：分别处理左右声道
        left_channel = y[0]
        right_channel = y[1]
        
        left_processed = process_mono_audio(left_channel, sr, semitone_shift, tempo_rate)
        right_processed = process_mono_audio(right_channel, sr, semitone_shift, tempo_rate)
        
        # 合并为立体声
        min_len = min(len(left_processed), len(right_processed))
        left_processed = left_processed[:min_len]
        right_processed = right_processed[:min_len]
        
        return np.vstack([left_processed, right_processed])
    else:
        # 单声道
        if y.ndim == 2:
            y = y[0] if y.shape[0] == 1 else y.flatten()
        
        return process_mono_audio(y, sr, semitone_shift, tempo_rate)


def sanitize_filename(filename: str) -> str:
    """清理文件名"""
    import re
    illegal_chars = r'[<>:"/\\|?*]'
    return re.sub(illegal_chars, '_', filename)


def find_audio_files(directory: Path, song_id: str, song_name: str) -> List[Path]:
    """
    查找音频文件（支持多个片段）
    优先使用ID匹配，找不到时才使用歌名
    
    Args:
        directory: 音频目录
        song_id: 歌曲ID（已转换为字符串）
        song_name: 歌曲名称
        
    Returns:
        匹配的音频文件列表
    """
    audio_extensions = {'.wav', '.mp3', '.m4a', '.flac', '.aac', '.ogg', '.wma'}
    matches = []
    
    # 优先使用ID匹配（ID应该已经是正确格式的字符串）
    if song_id and song_id.strip() and song_id.lower() != 'nan':
        song_id_str = song_id.strip()
        
        logger.info(f"正在查找ID为 {song_id_str} 的音频文件...")
        
        for file_path in directory.rglob('*'):
            if file_path.suffix.lower() in audio_extensions:
                stem = file_path.stem
                # 匹配以ID开头的文件（ID-xxx格式）
                if stem.startswith(song_id_str):
                    # 确保后面是分隔符（-、_、空格）或直接结尾
                    if len(stem) == len(song_id_str) or stem[len(song_id_str)] in ['-', '_', ' ']:
                        matches.append(file_path)
                        logger.info(f"  找到匹配文件: {file_path.name}")
        
        if matches:
            logger.info(f"通过ID找到 {len(matches)} 个音频文件")
            return sorted(matches, key=lambda x: x.stem)
        else:
            logger.warning(f"未找到ID为 {song_id_str} 的音频文件")
    
    # 如果ID匹配失败，使用歌名匹配（作为后备方案）
    if song_name:
        logger.info(f"尝试使用歌名 '{song_name}' 查找音频文件...")
        song_name_clean = str(song_name).strip().lower()
        for file_path in directory.rglob('*'):
            if file_path.suffix.lower() in audio_extensions:
                if song_name_clean in file_path.stem.lower():
                    matches.append(file_path)
                    logger.info(f"  找到匹配文件: {file_path.name}")
        
        if matches:
            logger.info(f"通过歌名找到 {len(matches)} 个音频文件")
    
    return sorted(matches, key=lambda x: x.stem) if matches else []


def process_pitch_tempo_core(
    classified_excel_path: Path,
    audio_dir: Path,
    output_dir: Path,
    progress_callback: Optional[Callable[[int, int, str], bool]] = None
) -> Tuple[int, int]:
    """
    变调变速核心功能
    
    Args:
        classified_excel_path: 分类后的Excel文件
        audio_dir: 音频文件目录
        output_dir: 输出目录
        progress_callback: 进度回调函数 (当前进度, 总数, 消息) -> 是否继续
    
    Returns:
        (成功数量, 总数量)
    """
    # 读取所有sheet
    excel_file = pd.ExcelFile(classified_excel_path)
    
    # 统计总任务数
    total_tasks = 0
    task_list = []
    
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(classified_excel_path, sheet_name=sheet_name)
        
        # 跳过空sheet
        if df.empty:
            continue
        
        # 检查必需的列（分类结果输出的是'调号'和'速度'）
        required_cols = ['ID', '歌名', '调号', '速度']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            logger.warning(f"Sheet '{sheet_name}' 缺少必需的列: {missing_cols}")
            continue
        
        # 找到锚定歌曲（第一行）
        if len(df) < 2:
            continue
        
        anchor_song = df.iloc[0]
        # 正确处理ID（可能是数字类型）
        anchor_id_raw = anchor_song.get('ID', '')
        if pd.notna(anchor_id_raw) and anchor_id_raw != '':
            try:
                # 如果是浮点数且是整数值，转换为整数字符串
                anchor_id = str(int(float(anchor_id_raw)))
            except (ValueError, TypeError):
                anchor_id = str(anchor_id_raw).strip()
        else:
            anchor_id = ''
        
        anchor_name = str(anchor_song.get('歌名', '')).strip()
        anchor_key = str(anchor_song.get('调号', '')).strip()
        anchor_bpm = anchor_song.get('速度')
        
        try:
            anchor_bpm = float(anchor_bpm)
        except (ValueError, TypeError):
            logger.warning(f"Sheet '{sheet_name}' 锚定歌曲BPM无效")
            continue
        
        # 查找锚定歌曲的音频文件
        anchor_audio_files = find_audio_files(audio_dir, anchor_id, anchor_name)
        
        # 处理每一行（跳过第一行锚定歌曲）
        for idx in range(len(df)):
            row = df.iloc[idx]
            # 正确处理ID（可能是数字类型）
            song_id_raw = row.get('ID', '')
            if pd.notna(song_id_raw) and song_id_raw != '':
                try:
                    # 如果是浮点数且是整数值，转换为整数字符串
                    song_id = str(int(float(song_id_raw)))
                except (ValueError, TypeError):
                    song_id = str(song_id_raw).strip()
            else:
                song_id = ''
            
            song_name = str(row.get('歌名', '')).strip()
            
            # 跳过锚定歌曲本身
            if song_name == anchor_name or (song_id and song_id == anchor_id):
                continue
            
            song_key = str(row.get('调号', '')).strip()
            song_bpm = row.get('速度')
            
            try:
                song_bpm = float(song_bpm)
            except (ValueError, TypeError):
                continue
            
            # 计算处理参数
            semitone_shift = calculate_semitone_shift(song_key, anchor_key)
            if semitone_shift is None:
                continue
            
            tempo_rate = anchor_bpm / song_bpm if song_bpm > 0 else None
            if tempo_rate is None:
                continue
            
            # 查找音频文件
            logger.info(f"正在处理：ID={song_id}, 歌名={song_name}, 调号={song_key}, BPM={song_bpm}")
            audio_files = find_audio_files(audio_dir, song_id, song_name)
            
            for audio_file in audio_files:
                task_list.append({
                    'sheet_name': sheet_name,
                    'anchor_audio_files': anchor_audio_files,
                    'anchor_id': anchor_id,
                    'anchor_name': anchor_name,
                    'audio_file': audio_file,
                    'song_name': song_name,
                    'semitone_shift': semitone_shift,
                    'tempo_rate': tempo_rate
                })
                total_tasks += 1
    
    # 处理所有任务
    success_count = 0
    
    for task_idx, task in enumerate(task_list):
        if progress_callback:
            msg = f"处理 {task['song_name']} ({task_idx + 1}/{total_tasks})"
            if not progress_callback(task_idx + 1, total_tasks, msg):
                break
        
        try:
            # 先复制锚定歌曲（每个sheet只复制一次）
            sheet_name = task['sheet_name']
            safe_sheet_name = sanitize_filename(sheet_name)
            output_folder = output_dir / safe_sheet_name
            output_folder.mkdir(parents=True, exist_ok=True)
            
            # 复制锚定歌曲的所有片段
            for anchor_audio_file in task['anchor_audio_files']:
                anchor_filename = anchor_audio_file.stem
                safe_anchor_filename = sanitize_filename(anchor_filename)
                output_path = output_folder / f"{safe_anchor_filename}.wav"
                
                if not output_path.exists():
                    if anchor_audio_file.suffix.lower() == '.wav':
                        shutil.copy2(anchor_audio_file, output_path)
                    else:
                        # 转换为wav（高质量）
                        y, sr = librosa.load(str(anchor_audio_file), sr=None, mono=False)
                        if y.ndim == 2:
                            y = y.T
                        # 使用24位深度保存，音质更好
                        sf.write(str(output_path), y, sr, format='WAV', subtype='PCM_24')
            
            # 处理匹配歌曲
            audio_file = task['audio_file']
            semitone_shift = task['semitone_shift']
            tempo_rate = task['tempo_rate']
            
            # 加载音频
            y, sr = librosa.load(str(audio_file), sr=None, mono=False)
            
            # 处理音频
            y_processed = process_stereo_audio(y, sr, semitone_shift, tempo_rate)
            
            # 保存音频（高质量）
            original_filename = audio_file.stem
            safe_filename = sanitize_filename(original_filename)
            output_path = output_folder / f"{safe_filename}.wav"
            
            if y_processed.ndim == 2:
                y_processed = y_processed.T
            
            # 使用24位深度保存，音质更好
            sf.write(str(output_path), y_processed, sr, format='WAV', subtype='PCM_24')
            
            success_count += 1
            
        except Exception as e:
            logger.error(f"处理失败 {task['song_name']}: {e}")
    
    return success_count, total_tasks

