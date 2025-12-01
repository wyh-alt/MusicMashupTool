"""
步骤1：歌曲分类核心功能
从原始的 song_classifier_gui.py 中提取
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
from typing import Callable, Optional, Tuple, List


# 调号映射：字母调号到数字的转换
KEY_MAPPING = {
    'C': 0, 'C#': 1, 'Db': 1,
    'D': 2, 'D#': 3, 'Eb': 3,
    'E': 4,
    'F': 5, 'F#': 6, 'Gb': 6,
    'G': 7, 'G#': 8, 'Ab': 8,
    'A': 9, 'A#': 10, 'Bb': 10,
    'B': 11
}

# 数字到字母调号的映射
NUM_TO_KEY = {
    0: 'C', 1: 'C#', 2: 'D', 3: 'D#', 4: 'E',
    5: 'F', 6: 'F#', 7: 'G', 8: 'G#', 9: 'A',
    10: 'A#', 11: 'B'
}


def key_to_number(key):
    """将调号转换为数字"""
    if pd.isna(key):
        raise ValueError("调号不能为空")
    
    key_str = str(key).strip()
    
    # 如果是数字，直接转换
    try:
        num = int(float(key_str))
        return num % 12
    except ValueError:
        pass
    
    # 尝试字母映射
    if key_str in KEY_MAPPING:
        return KEY_MAPPING[key_str]
    
    # 大小写不敏感匹配
    key_upper = key_str.upper()
    for k, v in KEY_MAPPING.items():
        if k.upper() == key_upper:
            return v
    
    raise ValueError(f"无法识别的调号: {key_str}")


def number_to_key(num):
    """将数字调号转换为字母调号"""
    num = int(num) % 12
    return NUM_TO_KEY.get(num, str(num))


def classify_songs_core(
    excel_path: Path,
    output_path: Path,
    progress_callback: Optional[Callable[[int, int, str], bool]] = None,
    key_range: int = 2,
    bpm_range: int = 5
) -> Tuple[List[List[int]], pd.DataFrame]:
    """
    歌曲分类核心功能
    
    Args:
        excel_path: 输入的Excel文件路径
        output_path: 输出的Excel文件路径
        progress_callback: 进度回调函数 (当前进度, 总数, 消息) -> 是否继续
        key_range: 调号匹配区间（±几个半音），默认2
        bpm_range: 速度匹配区间（±几个BPM），默认5
    
    Returns:
        (groups, df) - 分类组列表和原始数据DataFrame
    """
    # 读取Excel文件
    df = pd.read_excel(excel_path)
    
    # 检查必需的列
    has_song_name = '歌名' in df.columns or 'name' in df.columns
    has_key = '调号' in df.columns or 'key' in df.columns
    has_bpm = '速度' in df.columns or 'bpm' in df.columns
    
    if not (has_song_name and has_key and has_bpm):
        raise ValueError("Excel文件必须包含 '歌名'、'调号'、'速度' 三列")
    
    # 统一列名
    if 'name' in df.columns and '歌名' not in df.columns:
        df.rename(columns={'name': '歌名'}, inplace=True)
    if 'key' in df.columns and '调号' not in df.columns:
        df.rename(columns={'key': '调号'}, inplace=True)
    if 'bpm' in df.columns and '速度' not in df.columns:
        df.rename(columns={'bpm': '速度'}, inplace=True)
    
    # 处理可选列
    optional_columns = {
        'ID': 'id',
        'Chord Ai': 'Chord Ai',
        '歌手': '歌手',
        '副歌开始时间': '副歌开始时间',
        '副歌结束时间': '副歌结束时间',
        '段落剪切时间': '段落剪切时间',
        '性别': '性别'
    }
    
    # 统一列名
    for display_name, internal_name in optional_columns.items():
        if internal_name == 'id':
            if 'ID' in df.columns:
                df.rename(columns={'ID': 'id'}, inplace=True)
            elif 'id' not in df.columns:
                df['id'] = df.index + 1
        elif internal_name not in df.columns:
            df[internal_name] = ''
    
    # 处理性别列的不同写法
    gender_column = None
    for col_name in ['性别', 'gender', 'Gender', 'SEX', 'sex']:
        if col_name in df.columns:
            gender_column = col_name
            break
    if gender_column and gender_column != '性别':
        df.rename(columns={gender_column: '性别'}, inplace=True)
    elif '性别' not in df.columns:
        df['性别'] = ''
    
    # 处理歌手列
    if '歌手名' in df.columns:
        df.rename(columns={'歌手名': '歌手'}, inplace=True)
    elif 'artist' in df.columns:
        df.rename(columns={'artist': '歌手'}, inplace=True)
    elif '歌手' not in df.columns:
        df['歌手'] = ''
    
    # 规范化性别信息
    df['性别'] = df['性别'].fillna('').astype(str).str.strip()
    df['gender_normalized'] = df['性别'].str.lower()
    
    # 保存原始调号值
    df['调号_original'] = df['调号'].copy()
    
    # 转换调号为数字
    df['key_num'] = df['调号'].apply(key_to_number)
    
    # 转换速度为数值
    df['速度'] = pd.to_numeric(df['速度'], errors='coerce')
    
    if df['速度'].isna().any():
        raise ValueError("Excel文件中存在无效的速度值")
    
    # 分类处理
    groups = []
    used_as_anchor = set()
    
    total_songs = len(df)
    
    for anchor_idx in range(total_songs):
        if progress_callback:
            if not progress_callback(anchor_idx + 1, total_songs, f"分析第 {anchor_idx + 1}/{total_songs} 首歌曲"):
                return groups, df
        
        current_group = [anchor_idx]
        used_as_anchor.add(anchor_idx)
        
        # 获取锚点信息
        anchor_key_num = df.iloc[anchor_idx]['key_num']
        anchor_bpm = df.iloc[anchor_idx]['速度']
        anchor_gender = df.iloc[anchor_idx]['gender_normalized']
        
        # 查找匹配歌曲
        for candidate_idx in range(total_songs):
            if candidate_idx in used_as_anchor:
                continue
            
            candidate_key_num = df.iloc[candidate_idx]['key_num']
            candidate_bpm = df.iloc[candidate_idx]['速度']
            candidate_gender = df.iloc[candidate_idx]['gender_normalized']
            
            # 性别不同跳过
            if candidate_gender != anchor_gender:
                continue
            
            # 检查调号差异（使用用户设定的区间）
            key_diff = abs(candidate_key_num - anchor_key_num)
            key_diff_circular = min(key_diff, 12 - key_diff)
            if key_diff_circular > key_range:
                continue
            
            # 检查速度差异（使用用户设定的区间）
            bpm_diff = abs(candidate_bpm - anchor_bpm)
            if bpm_diff <= bpm_range:
                current_group.append(candidate_idx)
        
        groups.append(current_group)
    
    # 生成Excel文件
    wb = Workbook()
    wb.remove(wb.active)
    
    sheet_name_counts = {}
    
    for group_num, group_indices in enumerate(groups, start=1):
        if len(group_indices) == 1:
            continue
        
        group_songs = df.iloc[group_indices].copy()
        anchor_song = group_songs.iloc[0]
        
        # 生成sheet名称
        anchor_song_name = str(anchor_song['歌名']).strip()
        for ch in r'[]:*?/\\':
            anchor_song_name = anchor_song_name.replace(ch, '-')
        base_sheet_name = anchor_song_name.strip() or f"Group{group_num}"
        base_sheet_name = base_sheet_name[:31]
        
        # 处理重名
        sheet_name_count = sheet_name_counts.get(base_sheet_name, 0)
        if sheet_name_count == 0:
            sheet_name = base_sheet_name
        else:
            suffix = f"_{sheet_name_count}"
            available_length = 31 - len(suffix)
            sheet_name = (base_sheet_name[:available_length] + suffix)
        sheet_name_counts[base_sheet_name] = sheet_name_count + 1
        
        ws = wb.create_sheet(title=sheet_name)
        
        # 生成匹配歌曲列表
        match_songs = []
        for j in range(1, len(group_indices)):
            match_song = group_songs.iloc[j]
            combined_name = f"{anchor_song['歌名']}+{match_song['歌名']}"
            match_songs.append({
                'match_song': match_song,
                'combined_name': combined_name
            })
        
        # 定义列
        original_columns = ['ID', 'Chord Ai', '歌名', '歌手', '副歌开始时间', 
                          '副歌结束时间', '段落剪切时间', '调号', '速度', '性别']
        headers = original_columns + ['成品名']
        ws.append(headers)
        
        # 设置样式
        yellow_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
        green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        header_font = Font(bold=True)
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        num_original_cols = len(original_columns)
        for col_idx in range(1, num_original_cols + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = headers[col_idx - 1]
            cell.fill = yellow_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
        
        product_col_idx = num_original_cols + 1
        cell = ws.cell(row=1, column=product_col_idx)
        cell.value = headers[-1]
        cell.fill = green_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
        
        # 写入数据
        current_row = 2
        column_mapping = {
            'ID': 'id',
            'Chord Ai': 'Chord Ai',
            '歌名': '歌名',
            '歌手': '歌手',
            '副歌开始时间': '副歌开始时间',
            '副歌结束时间': '副歌结束时间',
            '段落剪切时间': '段落剪切时间',
            '调号': '调号_original',
            '速度': '速度',
            '性别': '性别'
        }
        
        for match_data in match_songs:
            match_song = match_data['match_song']
            combined_name = match_data['combined_name']
            
            # 第一行：锚定歌曲
            for col_idx, col_name in enumerate(original_columns, start=1):
                df_col_name = column_mapping.get(col_name, col_name)
                value = anchor_song.get(df_col_name, '')
                # 确保ID不是NaN，并转换为字符串或整数
                if col_name == 'ID' and value and not pd.isna(value):
                    try:
                        value = int(value) if float(value) == int(float(value)) else str(value)
                    except:
                        value = str(value)
                ws.cell(row=current_row, column=col_idx).value = value
            
            # 第二行：匹配歌曲
            for col_idx, col_name in enumerate(original_columns, start=1):
                df_col_name = column_mapping.get(col_name, col_name)
                value = match_song.get(df_col_name, '')
                # 确保ID不是NaN，并转换为字符串或整数
                if col_name == 'ID' and value and not pd.isna(value):
                    try:
                        value = int(value) if float(value) == int(float(value)) else str(value)
                    except:
                        value = str(value)
                ws.cell(row=current_row + 1, column=col_idx).value = value
            
            # 合并成品名列
            product_col_letter = get_column_letter(product_col_idx)
            ws.merge_cells(f'{product_col_letter}{current_row}:{product_col_letter}{current_row + 1}')
            ws.cell(row=current_row, column=product_col_idx).value = combined_name
            merged_cell = ws.cell(row=current_row, column=product_col_idx)
            merged_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            merged_cell.border = thin_border
            
            # 设置边框
            for row in [current_row, current_row + 1]:
                for col in range(1, product_col_idx + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                    cell.border = thin_border
            
            current_row += 2
        
        # 调整列宽
        column_widths = {
            'A': 12, 'B': 15, 'C': 25, 'D': 20, 'E': 15,
            'F': 15, 'G': 15, 'H': 10, 'I': 10, 'J': 10, 'K': 40
        }
        for col_letter, width in column_widths.items():
            ws.column_dimensions[col_letter].width = width
    
    # 保存文件
    wb.save(output_path)
    
    return groups, df

