"""
步骤3：音频拼接核心功能
从原始的音频拼接工具中提取
"""
import pandas as pd
import logging
from pathlib import Path
from typing import Callable, Optional, Tuple
from pydub import AudioSegment

logger = logging.getLogger(__name__)


def find_audio_file(audio_folder: Path, song_id: str, segment_type: str) -> Optional[Path]:
    """
    查找音频文件
    
    Args:
        audio_folder: 音频文件所在文件夹
        song_id: 歌曲ID
        segment_type: 片段类型（"前段副歌" 或 "后段副歌"）
        
    Returns:
        文件路径，如果未找到返回None
    """
    # 构建文件名模式
    filename = f"{song_id}-{segment_type}"
    
    # 支持的音频格式
    audio_extensions = ['.mp3', '.wav', '.m4a', '.flac', '.ogg', '.aac']
    
    # 在ID子文件夹中查找
    for ext in audio_extensions:
        file_path = audio_folder / song_id / f"{filename}{ext}"
        if file_path.exists():
            return file_path
    
    # 在根目录查找
    for ext in audio_extensions:
        file_path = audio_folder / f"{filename}{ext}"
        if file_path.exists():
            return file_path
    
    return None


def create_silence(duration_seconds: float) -> AudioSegment:
    """创建指定时长的静音"""
    if duration_seconds <= 0:
        return AudioSegment.empty()
    
    silence = AudioSegment.silent(duration=int(duration_seconds * 1000))
    return silence


def concat_audio_pair(
    audio_folder: Path,
    id_a: str,
    id_b: str,
    gap_duration: float
) -> AudioSegment:
    """
    拼接一对歌曲的音频
    
    拼接顺序：
    静音 -> A的前段副歌 -> 静音 -> B的前段副歌 -> 静音 -> A的后段副歌 -> 静音 -> B的后段副歌 -> 静音
    
    Args:
        audio_folder: 音频文件所在文件夹
        id_a: 第一个歌曲ID
        id_b: 第二个歌曲ID
        gap_duration: 静音间隙时长（秒）
        
    Returns:
        拼接后的音频
    """
    # 查找所有需要的音频文件
    audio_a_front_path = find_audio_file(audio_folder, id_a, "前段副歌")
    audio_b_front_path = find_audio_file(audio_folder, id_b, "前段副歌")
    audio_a_back_path = find_audio_file(audio_folder, id_a, "后段副歌")
    audio_b_back_path = find_audio_file(audio_folder, id_b, "后段副歌")
    
    if not audio_a_front_path:
        raise FileNotFoundError(f"找不到音频文件: {id_a}-前段副歌")
    if not audio_b_front_path:
        raise FileNotFoundError(f"找不到音频文件: {id_b}-前段副歌")
    if not audio_a_back_path:
        raise FileNotFoundError(f"找不到音频文件: {id_a}-后段副歌")
    if not audio_b_back_path:
        raise FileNotFoundError(f"找不到音频文件: {id_b}-后段副歌")
    
    # 加载音频
    audio_a_front = AudioSegment.from_file(str(audio_a_front_path))
    audio_b_front = AudioSegment.from_file(str(audio_b_front_path))
    audio_a_back = AudioSegment.from_file(str(audio_a_back_path))
    audio_b_back = AudioSegment.from_file(str(audio_b_back_path))
    
    # 创建静音
    silence = create_silence(gap_duration)
    
    # 按顺序拼接（开头和结尾都加静音）
    result = (
        silence + audio_a_front + 
        silence + audio_b_front + 
        silence + audio_a_back + 
        silence + audio_b_back + 
        silence
    )
    
    return result


def concat_audio_core(
    classified_excel_path: Path,
    processed_audio_dir: Path,
    output_dir: Path,
    gap_duration: float,
    progress_callback: Optional[Callable[[int, int, str], bool]] = None
) -> Tuple[int, int]:
    """
    音频拼接核心功能
    
    Args:
        classified_excel_path: 分类后的Excel文件
        processed_audio_dir: 处理后的音频目录
        output_dir: 输出目录
        gap_duration: 静音间隙时长（秒）
        progress_callback: 进度回调函数 (当前进度, 总数, 消息) -> 是否继续
    
    Returns:
        (成功数量, 总数量)
    """
    # 读取所有sheet
    excel_data = pd.read_excel(classified_excel_path, sheet_name=None)
    
    # 统计总任务数
    total_pairs = 0
    for sheet_name, df in excel_data.items():
        if 'ID' in df.columns:
            ids = df['ID'].dropna().astype(str).tolist()
            pair_count = len(ids) // 2
            total_pairs += pair_count
    
    success_count = 0
    current_pair = 0
    
    # 处理每个Sheet
    for sheet_name, df in excel_data.items():
        # 检查必需的列
        if 'ID' not in df.columns:
            logger.warning(f"Sheet '{sheet_name}' 缺少 'ID' 列，跳过")
            continue
        
        if '成品名' not in df.columns:
            logger.warning(f"Sheet '{sheet_name}' 缺少 '成品名' 列，跳过")
            continue
        
        # 获取ID列表
        ids = df['ID'].dropna().astype(str).tolist()
        output_names = df['成品名'].dropna().astype(str).tolist()
        
        if len(ids) == 0:
            logger.warning(f"Sheet '{sheet_name}' 没有有效的ID数据")
            continue
        
        if len(ids) % 2 != 0:
            logger.warning(f"Sheet '{sheet_name}' ID数量为奇数（{len(ids)}），最后一个ID将被忽略")
        
        # 音频文件夹路径（假设文件夹名称与Sheet名称相同）
        audio_folder = processed_audio_dir / sheet_name
        if not audio_folder.exists():
            logger.error(f"音频文件夹不存在: {audio_folder}")
            continue
        
        # 创建输出文件夹
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 两两配对处理
        pair_count = len(ids) // 2
        
        for i in range(pair_count):
            id_a = ids[i * 2]
            id_b = ids[i * 2 + 1]
            
            current_pair += 1
            
            # 获取输出文件名
            if i < len(output_names):
                output_name = output_names[i]
            else:
                output_name = f"{id_a}_{id_b}"
            
            if progress_callback:
                msg = f"拼接 {output_name} ({current_pair}/{total_pairs})"
                if not progress_callback(current_pair, total_pairs, msg):
                    return success_count, total_pairs
            
            try:
                # 拼接音频
                result_audio = concat_audio_pair(audio_folder, id_a, id_b, gap_duration)
                
                # 保存文件
                output_path = output_dir / f"{output_name}.mp3"
                result_audio.export(str(output_path), format="mp3", bitrate="320k")
                
                logger.info(f"已保存: {output_path}")
                success_count += 1
                
            except Exception as e:
                logger.error(f"拼接失败 {output_name}: {e}")
    
    return success_count, total_pairs


