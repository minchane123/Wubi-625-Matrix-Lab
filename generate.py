# -*- coding: utf-8 -*-
import os
import pandas as pd
import string
from openpyxl.styles import PatternFill, Alignment
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.utils import get_column_letter

def load_freq_and_level(file_path):
    char_rank = {}
    tier_2500 = set()
    tier_3500 = set()
    tier_5000 = set()
    if not os.path.exists(file_path):
        return char_rank, tier_2500, tier_3500, tier_5000

    df = pd.read_csv(file_path)
    for _, row in df.iterrows():
        char = row['character']
        try:
            rank = int(row['frequency_rank'])
        except:
            rank = 9999999
        char_rank[char] = rank
        if 1 <= rank <= 2500:
            tier_2500.add(char)
        elif 2500 < rank <= 3500:
            tier_3500.add(char)
        elif 3500 < rank <= 5000:
            tier_5000.add(char)
    return char_rank, tier_2500, tier_3500, tier_5000

def generate_matrix():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(current_dir, "data")
    output_dir = os.path.join(current_dir, "output")
    os.makedirs(data_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    csv_path = os.path.join(data_dir, "系统码表.csv")
    excel_path = os.path.join(output_dir, "五笔625炼药台.xlsx")
    hanzidb_path = os.path.join(data_dir, "hanzi_db.csv")

    char_rank, TIER_2500, TIER_3500, TIER_5000 = load_freq_and_level(hanzidb_path)

    if not os.path.exists(csv_path):
        return

    dict_2char = {}  
    dict_prefix = {} 
    encodings = ['utf-8', 'gbk', 'utf-8-sig']
    for enc in encodings:
        try:
            with open(csv_path, 'r', encoding=enc) as f:
                lines = f.readlines()
            break
        except UnicodeDecodeError:
            continue

    valid_letters = set(string.ascii_lowercase[:25])
    line_idx = 0
    for line in lines:
        line_idx += 1
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        parts = line.replace('\t', ',').replace(' ', ',').split(',')
        parts = [p.strip() for p in parts if p.strip()]
        if len(parts) < 2:
            continue
        code = parts[0].lower()
        if 2 <= len(code) <= 4 and all(c in valid_letters for c in code):
            single_chars = [word for word in parts[1:] if len(word) == 1]
            if not single_chars:
                continue
            prefix = code[:2]
            suffix = code[2:] 
            if len(code) == 2:
                if code not in dict_2char:
                    dict_2char[code] = single_chars[0]
            elif len(code) in (3, 4):
                if prefix not in dict_prefix:
                    dict_prefix[prefix] = []
                for char in single_chars:
                    if not any(item[0] == char for item in dict_prefix[prefix]):
                        dict_prefix[prefix].append((char, suffix, line_idx))

    sheet1_chars_set = set(dict_2char.values())

    final_prefix_data = {}
    for prefix, items in dict_prefix.items():
        valid_items = []
        for char, suffix, line_num in items:
            rank = char_rank.get(char, 9999999)
            if rank <= 6000:
                valid_items.append((char, suffix, rank))
        valid_items.sort(key=lambda x: x[2])
        final_prefix_data[prefix] = []
        for char, suffix, _ in valid_items:
            if char in sheet1_chars_set:
                suffix = ""
            final_prefix_data[prefix].append((char, suffix))

    letters = [c for c in string.ascii_lowercase if c != 'z']
    level_1_simple = {
        'a': '工', 'b': '了', 'c': '以', 'd': '在', 'e': '有', 'f': '地', 'g': '一',
        'h': '上', 'i': '不', 'j': '是', 'k': '中', 'l': '国', 'm': '同', 'n': '民',
        'o': '为', 'p': '这', 'q': '我', 'r': '的', 's': '要', 't': '和', 'u': '产',
        'v': '发', 'w': '人', 'x': '经', 'y': '主'
    }

    matrix_data = []
    for first_char in letters:
        row = {"首码": first_char.upper()}
        char_l1 = level_1_simple.get(first_char, "")
        if char_l1:
            rank = char_rank.get(char_l1, 9999999)
            rank_str = str(rank).zfill(4) if rank < 9999999 else "----"
            row["一级简码"] = f"{char_l1}|{rank_str}"
        else:
            row["一级简码"] = ""

        for second_char in letters:
            code = first_char + second_char
            char = dict_2char.get(code, "")
            if char:
                rank = char_rank.get(char, 9999999)
                rank_str = str(rank).zfill(4) if rank < 9999999 else "----"
                row[second_char.upper()] = f"{char}|{rank_str}"
            else:
                row[second_char.upper()] = ""
        matrix_data.append(row)

    df_2char = pd.DataFrame(matrix_data)
    col_headers = ['首码', '一级简码'] + [c.upper() for c in letters]
    df_2char = df_2char[col_headers]

    list_prefix_rows = []
    merge_info = [] 
    current_excel_row = 2 
    for first_char in letters:
        for second_char in letters:
            prefix = first_char + second_char
            chars = final_prefix_data.get(prefix, [])
            if not chars:
                row = {"前缀编码": prefix.upper()}
                for i in range(10): row[f"候选{i+1}"] = ""
                list_prefix_rows.append(row)
                current_excel_row += 1
            else:
                chunks = [chars[i:i + 10] for i in range(0, len(chars), 10)]
                start_row = current_excel_row
                for chunk in chunks:
                    row = {"前缀编码": prefix.upper()}
                    for i in range(10):
                        if i < len(chunk):
                            char, suffix = chunk[i]
                            rank = char_rank.get(char, 9999999)
                            rank_str = str(rank).zfill(4) if rank < 9999999 else "----"
                            row[f"候选{i+1}"] = f"{char}{suffix}|{rank_str}"
                        else:
                            row[f"候选{i+1}"] = ""
                    list_prefix_rows.append(row)
                    current_excel_row += 1
                end_row = current_excel_row - 1
                if end_row > start_row:
                    merge_info.append((start_row, end_row))

    df_prefix_all = pd.DataFrame(list_prefix_rows)

    try:
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df_2char.to_excel(writer, sheet_name='二级简码矩阵', index=False)
            df_prefix_all.to_excel(writer, sheet_name='词频同前缀候选', index=False)

            workbook = writer.book
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
            orange_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

            # === 核心：使用左对齐，彻底摒弃分散对齐 ===
            left_align = Alignment(horizontal='left', vertical='center', indent=0)
            center_align = Alignment(horizontal='center', vertical='center')

            # === 核心：强制使用“新宋体”等宽字体，确保 1空格 = 1字母，1汉字 = 2字母 ===
            font_char = InlineFont(rFont='新宋体', sz=11)
            font_rank = InlineFont(rFont='新宋体', sz=8, color='808080')

            def apply_rich_text_and_style(cell):
                if not cell.value: return
                val_str = str(cell.value)
                if '|' not in val_str: return

                char_part, rank_part = val_str.split('|')

                # 计算“字+码”的视觉宽度（汉字=2，字母=1）
                visual_width = sum(2 if '\u4e00' <= c <= '\u9fa5' else 1 for c in char_part)
                
                # 目标宽度设为 4（足以容纳所有 1字+2码 的组合）
                # 缺多少用半角空格补多少
                spaces = " " * max(0, 4 - visual_width)
                
                # 拼接：左侧（字+码+补齐空格） + 右侧（数字）
                cell.value = CellRichText(
                    TextBlock(font_char, f"{char_part}{spaces}"),
                    TextBlock(font_rank, rank_part)
                )

                char = char_part[0]
                if char in TIER_2500: cell.fill = green_fill
                elif char in TIER_3500: cell.fill = blue_fill
                elif char in TIER_5000: cell.fill = orange_fill

                cell.alignment = left_align

            # 处理 Sheet1
            ws1 = workbook['二级简码矩阵']
            ws1.column_dimensions['A'].width = 10
            ws1.column_dimensions['B'].width = 12
            for col_letter in letters:
                ws1.column_dimensions[col_letter.upper()].width = 11
            for row in ws1.iter_rows(min_row=1, max_row=ws1.max_row, min_col=1, max_col=1):
                for cell in row:
                    cell.font = cell.font.copy(bold=True)
            for row in ws1.iter_rows(min_row=2, min_col=2, max_col=len(col_headers)):
                for cell in row:
                    apply_rich_text_and_style(cell)

            # 处理 Sheet2
            ws2 = workbook['词频同前缀候选']
            ws2.column_dimensions['A'].width = 12
            for col_idx in range(2, 12):
                col_letter = get_column_letter(col_idx)
                ws2.column_dimensions[col_letter].width = 12
            for row in ws2.iter_rows(min_row=2, min_col=2, max_col=11):
                for cell in row:
                    apply_rich_text_and_style(cell)

            # 合并单元格
            for start_row, end_row in merge_info:
                ws2.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
                top_cell = ws2.cell(row=start_row, column=1)
                top_cell.alignment = center_align
                top_cell.font = top_cell.font.copy(bold=True) 
            for row in ws2.iter_rows(min_row=2, min_col=1, max_col=1):
                for cell in row:
                    cell.alignment = center_align
                    cell.font = cell.font.copy(bold=True)

        print(f"🎉 完美！《五笔625炼药台》已生成：\n👉 {excel_path}")
    except PermissionError:
        print(f"❌ 导出失败：请先关闭已经打开的 {excel_path} 文件！")
    except Exception as e:
        print(f"❌ 发生错误：{e}")

if __name__ == "__main__":
    generate_matrix()