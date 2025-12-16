"""
步骤3：音频拼接核心功能
从原始的音频拼接工具中提取
"""
import pandas as pd
import logging
import re
from pathlib import Path
from typing import Callable, Optional, Tuple, List
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
    
    # 在根目录查找（处理后的音频直接在输出文件夹中）
    for ext in audio_extensions:
        file_path = audio_folder / f"{filename}{ext}"
        if file_path.exists():
            return file_path
    
    # 在ID子文件夹中查找（兼容旧结构）
    for ext in audio_extensions:
        file_path = audio_folder / song_id / f"{filename}{ext}"
        if file_path.exists():
            return file_path
    
    return None


def parse_product_name(product_name: str) -> Tuple[str, str]:
    """
    从成品名解析出两个ID
    
    成品名格式：ID1-ID2-拼接成品
    例如：1086360-871604-拼接成品
    
    Args:
        product_name: 成品名字符串
        
    Returns:
        (id_a, id_b) 元组
    """
    # 匹配格式：数字-数字-其他内容
    match = re.match(r'^(\d+)-(\d+)-', product_name)
    if match:
        return match.group(1), match.group(2)
    
    # 尝试用'-'分割
    parts = product_name.split('-')
    if len(parts) >= 2:
        return parts[0], parts[1]
    
    return '', ''


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


def sanitize_filename(filename: str) -> str:
    """清理文件名，移除非法字符"""
    illegal_chars = r'[<>:"/\\|?*]'
    return re.sub(illegal_chars, '_', filename)


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
    # 读取分类结果（单一sheet）
    df = pd.read_excel(classified_excel_path, sheet_name=0)
    
    # 检查必需的列
    if 'ID' not in df.columns:
        logger.warning("分类结果缺少 'ID' 列")
        return 0, 0
    
    if '成品名' not in df.columns:
        logger.warning("分类结果缺少 '成品名' 列")
        return 0, 0
    
    # 新结构：每两行是一对（锚定+匹配）
    # 成品名格式：ID1-ID2-拼接成品
    num_pairs = len(df) // 2
    
    if num_pairs == 0:
        logger.warning("分类结果中没有有效的配对数据")
        return 0, 0
    
    success_count = 0
    output_dir.mkdir(parents=True, exist_ok=True)
    
    for pair_idx in range(num_pairs):
        row_anchor = pair_idx * 2  # 锚定歌曲行
        
        anchor_row = df.iloc[row_anchor]
        
        # 获取成品名
        product_name = str(anchor_row.get('成品名', '')).strip()
        if not product_name:
            logger.warning(f"配对 {pair_idx + 1} 缺少成品名，跳过")
            continue
        
        # 从成品名解析ID
        id_a, id_b = parse_product_name(product_name)
        if not id_a or not id_b:
            logger.warning(f"无法从成品名 '{product_name}' 解析ID，跳过")
            continue
        
        if progress_callback:
            msg = f"拼接 {product_name} ({pair_idx + 1}/{num_pairs})"
            if not progress_callback(pair_idx + 1, num_pairs, msg):
                return success_count, num_pairs
        
        try:
            # 音频文件夹路径（使用成品名作为子文件夹）
            safe_product_name = sanitize_filename(product_name)
            audio_folder = processed_audio_dir / safe_product_name
            
            if not audio_folder.exists():
                logger.error(f"音频文件夹不存在: {audio_folder}")
                continue
            
            # 拼接音频
            result_audio = concat_audio_pair(audio_folder, id_a, id_b, gap_duration)
            
            # 保存文件
            output_path = output_dir / f"{safe_product_name}.mp3"
            result_audio.export(str(output_path), format="mp3", bitrate="320k")
            
            logger.info(f"已保存: {output_path}")
            success_count += 1
            
        except Exception as e:
            logger.error(f"拼接失败 {product_name}: {e}")
            import traceback
            logger.error(traceback.format_exc())
    
    return success_count, num_pairs


