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

# 尝试导入 pyrubberband 并验证 rubberband-cli 是否可用
HAS_RUBBERBAND = False
RUBBERBAND_VERSION = None
prb = None

def _check_rubberband_cli():
    """检查 rubberband-cli 是否已安装且可用"""
    try:
        # Windows 下需要隐藏窗口
        startupinfo = None
        creationflags = 0
        if platform.system() == 'Windows':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            creationflags = subprocess.CREATE_NO_WINDOW
        
        result = subprocess.run(
            ['rubberband', '--version'],
            capture_output=True,
            text=True,
            timeout=5,
            startupinfo=startupinfo,
            creationflags=creationflags
        )
        if result.returncode == 0:
            version = result.stdout.strip() or result.stderr.strip()
            return True, version
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError) as e:
        logger.debug(f"rubberband-cli 检查失败: {e}")
    return False, None

try:
    import pyrubberband as prb
    # 仅导入成功不够，还需要验证 rubberband-cli 是否可用
    cli_available, cli_version = _check_rubberband_cli()
    if cli_available:
        HAS_RUBBERBAND = True
        RUBBERBAND_VERSION = cli_version
        logger.info(f"✓ pyrubberband 可用，rubberband-cli 版本: {cli_version}")
    else:
        logger.warning("⚠️ pyrubberband 已安装，但 rubberband-cli 不可用")
        logger.warning("⚠️ 将回退到 librosa 处理（音质较差）")
        logger.warning("⚠️ 请安装 rubberband-cli: conda install -c conda-forge rubberband")
except ImportError:
    logger.warning("⚠️ 未安装 pyrubberband - 将使用 librosa（音质较差）")
    logger.warning("⚠️ 强烈建议安装 pyrubberband 以获得最佳音质！")
    logger.warning("⚠️ 安装方法: pip install pyrubberband")


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
    # 记录处理参数（便于调试）
    logger.debug(f"[变调变速] 参数: semitone_shift={semitone_shift}, tempo_rate={tempo_rate:.4f}, sr={sr}, HAS_RUBBERBAND={HAS_RUBBERBAND}")
    
    # 如果同时需要变调和变速，使用 pyrubberband 一次性处理（音质最佳）
    if semitone_shift != 0 and tempo_rate != 1.0 and HAS_RUBBERBAND:
        try:
            # pyrubberband 同时处理变调和变速，避免二次处理损失
            logger.info(f"[pyrubberband] 变调 {semitone_shift:+d} 半音")
            y_processed = prb.pitch_shift(y, sr, semitone_shift)
            logger.info(f"[pyrubberband] 变速 {tempo_rate:.4f}x")
            y_processed = prb.time_stretch(y_processed, sr, tempo_rate)
            logger.info("[pyrubberband] 处理完成")
            return y_processed
        except Exception as e:
            logger.warning(f"pyrubberband 同时处理失败，使用分步处理: {e}")
    
    # 分步处理：先做升降调
    if semitone_shift != 0:
        if HAS_RUBBERBAND:
            try:
                # pyrubberband 的质量远高于 librosa
                logger.info(f"[pyrubberband] 变调 {semitone_shift:+d} 半音")
                y_shifted = prb.pitch_shift(y, sr, semitone_shift)
                logger.info("[pyrubberband] 变调完成")
            except Exception as e:
                logger.warning(f"pyrubberband 升降调失败，使用 librosa 高质量模式: {e}")
                # librosa 使用最高质量参数
                logger.info(f"[librosa] 变调 {semitone_shift:+d} 半音 (回退模式)")
                y_shifted = librosa.effects.pitch_shift(
                    y, 
                    sr=sr, 
                    n_steps=semitone_shift,
                    bins_per_octave=24,      # 提高音高分辨率（默认12）
                    res_type='soxr_vhq'      # 使用最高质量重采样器
                )
        else:
            # librosa 高质量模式
            logger.info(f"[librosa] 变调 {semitone_shift:+d} 半音")
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
                logger.info(f"[pyrubberband] 变速 {tempo_rate:.4f}x")
                y_final = prb.time_stretch(y_shifted, sr, tempo_rate)
                logger.info("[pyrubberband] 变速完成")
            except Exception as e:
                logger.warning(f"pyrubberband 变速失败，使用 librosa 高质量模式: {e}")
                # librosa 使用 phase vocoder，质量较好
                logger.info(f"[librosa] 变速 {tempo_rate:.4f}x (回退模式)")
                y_final = librosa.effects.time_stretch(
                    y_shifted, 
                    rate=tempo_rate
                )
        else:
            logger.info(f"[librosa] 变速 {tempo_rate:.4f}x")
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


def get_id_value(id_raw) -> str:
    """从原始ID值获取字符串形式的ID"""
    if pd.notna(id_raw) and id_raw != '':
        try:
            return str(int(float(id_raw)))
        except (ValueError, TypeError):
            return str(id_raw).strip()
    return ''


def get_audio_engine_info() -> str:
    """获取当前音频处理引擎信息"""
    if HAS_RUBBERBAND:
        return f"pyrubberband + rubberband-cli ({RUBBERBAND_VERSION or 'unknown version'})"
    else:
        return "librosa (回退模式，音质较差)"


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
    # 输出音频处理引擎信息（便于诊断不同设备上的问题）
    engine_info = get_audio_engine_info()
    logger.info(f"========== 音频处理引擎: {engine_info} ==========")
    
    # 读取分类结果（单一sheet）
    df = pd.read_excel(classified_excel_path, sheet_name=0)
    
    # 跳过空表格
    if df.empty:
        logger.warning("分类结果表格为空")
        return 0, 0
    
    # 检查必需的列
    required_cols = ['ID', '歌名', '调号', '速度', '成品名']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        logger.warning(f"分类结果缺少必需的列: {missing_cols}")
        return 0, 0
    
    # 统计总任务数和生成任务列表
    # 新结构：每两行是一对（锚定+匹配），成品名在合并单元格中
    total_tasks = 0
    task_list = []
    
    # 每两行为一对处理
    num_pairs = len(df) // 2
    
    for pair_idx in range(num_pairs):
        row_anchor = pair_idx * 2  # 锚定歌曲行
        row_match = pair_idx * 2 + 1  # 匹配歌曲行
        
        anchor_song = df.iloc[row_anchor]
        match_song = df.iloc[row_match]
        
        # 获取成品名（用于创建输出子文件夹）
        product_name = str(anchor_song.get('成品名', '')).strip()
        if not product_name:
            product_name = f"pair_{pair_idx + 1}"
        
        # 获取锚定歌曲信息
        anchor_id = get_id_value(anchor_song.get('ID', ''))
        anchor_name = str(anchor_song.get('歌名', '')).strip()
        anchor_key = str(anchor_song.get('调号', '')).strip()
        anchor_bpm = anchor_song.get('速度')
        
        try:
            anchor_bpm = float(anchor_bpm)
        except (ValueError, TypeError):
            logger.warning(f"配对 {pair_idx + 1} 锚定歌曲BPM无效")
            continue
        
        # 获取匹配歌曲信息
        match_id = get_id_value(match_song.get('ID', ''))
        match_name = str(match_song.get('歌名', '')).strip()
        match_key = str(match_song.get('调号', '')).strip()
        match_bpm = match_song.get('速度')
        
        try:
            match_bpm = float(match_bpm)
        except (ValueError, TypeError):
            logger.warning(f"配对 {pair_idx + 1} 匹配歌曲BPM无效")
            continue
        
        # 查找锚定歌曲的音频文件
        anchor_audio_files = find_audio_files(audio_dir, anchor_id, anchor_name)
        
        # 计算匹配歌曲的处理参数（将匹配歌曲调整到锚定歌曲的调和速度）
        semitone_shift = calculate_semitone_shift(match_key, anchor_key)
        if semitone_shift is None:
            logger.warning(f"无法计算调号差异: {match_key} -> {anchor_key}")
            continue
        
        tempo_rate = anchor_bpm / match_bpm if match_bpm > 0 else None
        if tempo_rate is None:
            logger.warning(f"无法计算速度比率: {match_bpm} -> {anchor_bpm}")
            continue
        
        # 详细记录变调变速参数（便于调试）
        logger.info(f"[参数计算] {match_name}: 调号 {match_key}->{anchor_key} = {semitone_shift:+d}半音, "
                   f"速度 {match_bpm:.1f}->{anchor_bpm:.1f} = {tempo_rate:.4f}x")
        
        # 查找匹配歌曲的音频文件
        logger.info(f"正在处理配对 {pair_idx + 1}: {anchor_name} + {match_name}")
        match_audio_files = find_audio_files(audio_dir, match_id, match_name)
        
        for audio_file in match_audio_files:
            task_list.append({
                'product_name': product_name,
                'anchor_audio_files': anchor_audio_files,
                'anchor_id': anchor_id,
                'anchor_name': anchor_name,
                'audio_file': audio_file,
                'match_id': match_id,
                'match_name': match_name,
                'semitone_shift': semitone_shift,
                'tempo_rate': tempo_rate
            })
            total_tasks += 1
    
    # 处理所有任务
    success_count = 0
    processed_files = set()  # 记录已处理的文件（product_name + audio_file_stem）
    copied_anchors = set()  # 记录已复制锚定歌曲的成品
    
    for task_idx, task in enumerate(task_list):
        if progress_callback:
            msg = f"处理 {task['match_name']} ({task_idx + 1}/{total_tasks})"
            if not progress_callback(task_idx + 1, total_tasks, msg):
                break
        
        try:
            product_name = task['product_name']
            safe_product_name = sanitize_filename(product_name)
            output_folder = output_dir / safe_product_name
            output_folder.mkdir(parents=True, exist_ok=True)
            
            # 先复制锚定歌曲（每个成品只复制一次）
            if product_name not in copied_anchors:
                for anchor_audio_file in task['anchor_audio_files']:
                    anchor_filename = anchor_audio_file.stem
                    safe_anchor_filename = sanitize_filename(anchor_filename)
                    output_path = output_folder / f"{safe_anchor_filename}.wav"
                    
                    if not output_path.exists():
                        if anchor_audio_file.suffix.lower() == '.wav':
                            shutil.copy2(anchor_audio_file, output_path)
                            logger.info(f"复制锚定歌曲: {output_path.name}")
                        else:
                            # 转换为wav（高质量）
                            y, sr = librosa.load(str(anchor_audio_file), sr=None, mono=False)
                            if y.ndim == 2:
                                y = y.T
                            # 使用24位深度保存，音质更好
                            sf.write(str(output_path), y, sr, format='WAV', subtype='PCM_24')
                            logger.info(f"转换并保存锚定歌曲: {output_path.name}")
                
                copied_anchors.add(product_name)
            
            # 处理匹配歌曲 - 检查是否已处理过
            audio_file = task['audio_file']
            file_key = f"{product_name}||{audio_file.stem}"
            
            if file_key in processed_files:
                logger.info(f"跳过重复处理: {audio_file.stem} (成品: {product_name})")
                success_count += 1
                continue
            
            semitone_shift = task['semitone_shift']
            tempo_rate = task['tempo_rate']
            
            logger.info(f"开始处理音频: {audio_file.name}, 变调={semitone_shift}半音, 变速={tempo_rate:.2f}x")
            
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
            logger.info(f"保存处理后音频: {output_path.name}")
            
            # 标记为已处理
            processed_files.add(file_key)
            success_count += 1
            
        except Exception as e:
            logger.error(f"处理失败 {task['match_name']}: {e}")
            import traceback
            logger.error(traceback.format_exc())
    
    return success_count, total_tasks

