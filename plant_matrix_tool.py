import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import os
import re
import sys
import os

# 确保打包后也能找到依赖
if hasattr(sys, '_MEIPASS'):
    # 打包后的运行环境
    os.chdir(sys._MEIPASS)

def natural_sort_key(s):
    """自然排序键函数，用于正确排序数字"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]


def process_excel_file():
    """处理Excel格式的植物样方数据，并按照样方编号排序"""
    file_path = filedialog.askopenfilename(
        title="选择Excel文件",
        filetypes=[("Excel文件", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    try:
        # 读取Excel文件
        df = pd.read_excel(file_path, header=None, engine='openpyxl')

        print(f"原始数据形状: {df.shape}")

        # 查找所有包含"物种"的单元格，这些是表格的起始位置
        species_positions = []
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell_value = str(df.iloc[i, j])
                if '物种' in cell_value and cell_value.strip() == '物种':
                    species_positions.append((i, j))

        print(f"找到 {len(species_positions)} 个表格起始位置: {species_positions}")

        if not species_positions:
            messagebox.showerror("错误", "未找到包含'物种'的表头，请检查文件格式")
            return

        # 处理每个表格
        all_tables_data = {}

        for idx, (start_row, start_col) in enumerate(species_positions):
            print(f"处理表格 {idx + 1}, 起始位置: ({start_row}, {start_col})")

            # 确定表格的结束位置（下一个表格开始或数据结束）
            end_row = len(df)
            if idx + 1 < len(species_positions):
                end_row = species_positions[idx + 1][0]

            # 提取表头
            header_row = df.iloc[start_row, start_col:]
            headers = []
            for j in range(len(header_row)):
                cell_val = str(header_row.iloc[j])
                if pd.isna(header_row.iloc[j]) or cell_val == 'nan' or not cell_val.strip():
                    break
                headers.append(cell_val)

            print(f"表格 {idx + 1} 表头: {headers}")

            # 提取数据行
            table_data = {}
            for i in range(start_row + 1, end_row):
                # 检查是否为空行
                if df.iloc[i].isna().all() or (df.iloc[i].astype(str) == '').all():
                    continue

                # 获取物种名称（第一个单元格）
                species_cell = str(df.iloc[i, start_col])
                if not species_cell or species_cell == 'nan' or species_cell.strip() == '':
                    continue

                # 提取物种名称（去除可能的数值）
                species_name = species_cell.split()[0] if ' ' in species_cell else species_cell

                # 提取该行的数值数据
                values = []
                for j in range(start_col + 1, start_col + len(headers)):
                    if j >= len(df.columns):
                        values.append(0)
                        continue

                    cell_val = df.iloc[i, j]
                    if pd.isna(cell_val) or str(cell_val).strip() == '':
                        values.append(0)
                    else:
                        try:
                            # 尝试转换为数值
                            num_val = float(cell_val)
                            values.append(int(num_val) if num_val.is_integer() else num_val)
                        except (ValueError, TypeError):
                            values.append(0)

                # 确保数值数量与表头数量匹配
                expected_values = len(headers) - 1  # 减去物种列
                if len(values) < expected_values:
                    values.extend([0] * (expected_values - len(values)))
                elif len(values) > expected_values:
                    values = values[:expected_values]

                table_data[species_name] = values

            # 存储表格数据
            if table_data:
                all_tables_data[f'Table_{idx + 1}'] = {
                    'headers': headers,
                    'data': table_data
                }

        # 合并所有表格
        if not all_tables_data:
            messagebox.showerror("错误", "未能提取到有效数据")
            return

        # 收集所有物种
        all_species = set()
        for table_info in all_tables_data.values():
            all_species.update(table_info['data'].keys())

        # 收集所有样方编号
        all_quadrats = set()
        for table_info in all_tables_data.values():
            # 跳过物种列
            for header in table_info['headers'][1:]:
                all_quadrats.add(header)

        # 按照自然顺序排序样方编号
        sorted_quadrats = sorted(all_quadrats, key=natural_sort_key)
        print(f"排序后的样方编号: {sorted_quadrats}")

        # 创建合并后的数据
        merged_data = []
        for species in sorted(all_species):
            row = {'物种': species}

            # 为每个样方编号添加数据
            for quadrat in sorted_quadrats:
                # 在所有表格中查找该样方编号的数据
                found_value = 0
                for table_info in all_tables_data.values():
                    headers = table_info['headers'][1:]  # 排除物种列
                    if quadrat in headers:
                        quadrat_index = headers.index(quadrat)
                        species_data = table_info['data'].get(species, [0] * len(headers))
                        if quadrat_index < len(species_data):
                            found_value = species_data[quadrat_index]
                            break  # 找到后跳出循环

                row[quadrat] = found_value

            merged_data.append(row)

        # 创建DataFrame
        result_df = pd.DataFrame(merged_data)

        # 重新排列列，使物种列在最前面，然后是排序后的样方编号
        cols = ['物种'] + sorted_quadrats
        result_df = result_df[cols]

        # 保存结果
        output_path = os.path.splitext(file_path)[0] + "_植物矩阵.xlsx"

        # 避免文件覆盖
        counter = 1
        original_output = output_path
        while os.path.exists(output_path):
            output_path = f"{os.path.splitext(original_output)[0]}_{counter}.xlsx"
            counter += 1

        result_df.to_excel(output_path, index=False, engine='openpyxl')

        # 显示成功信息
        messagebox.showinfo("处理完成",
                            f"✅ Excel文件处理成功！\n"
                            f"识别到 {len(all_tables_data)} 个表格\n"
                            f"合并为 {len(result_df)} 个物种\n"
                            f"输出 {len(sorted_quadrats)} 个样方\n"
                            f"输出文件：{os.path.basename(output_path)}")

        # 在控制台显示处理摘要
        print(f"\n处理摘要:")
        print(f"- 输入文件: {os.path.basename(file_path)}")
        print(f"- 识别表格: {len(all_tables_data)} 个")
        print(f"- 物种数量: {len(result_df)} 个")
        print(f"- 样方数量: {len(sorted_quadrats)} 个")
        print(f"- 输出文件: {output_path}")

    except Exception as e:
        error_msg = f"处理Excel文件时出错：\n{str(e)}"
        print(error_msg)
        messagebox.showerror("处理错误", error_msg)


def debug_excel_structure():
    """调试函数：显示Excel文件结构"""
    file_path = filedialog.askopenfilename(
        title="选择Excel文件进行调试",
        filetypes=[("Excel文件", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    try:
        # 读取Excel文件
        df = pd.read_excel(file_path, header=None, engine='openpyxl')

        debug_info = f"Excel文件结构分析:\n"
        debug_info += f"文件: {os.path.basename(file_path)}\n"
        debug_info += f"数据形状: {df.shape} (行×列)\n\n"

        # 查找所有包含"物种"的单元格
        species_cells = []
        for i in range(min(50, len(df))):  # 只检查前50行
            for j in range(min(20, len(df.columns))):  # 只检查前20列
                cell_value = str(df.iloc[i, j])
                if '物种' in cell_value:
                    species_cells.append((i, j, cell_value))

        debug_info += f"找到 {len(species_cells)} 个包含'物种'的单元格:\n"
        for i, j, value in species_cells:
            debug_info += f"  位置: ({i}, {j}), 值: '{value}'\n"

        debug_info += f"\n前10行数据预览:\n"
        debug_info += df.head(10).to_string() + "\n"

        # 显示调试窗口
        debug_window = tk.Toplevel()
        debug_window.title("Excel文件结构分析")
        debug_window.geometry("900x700")

        text_widget = tk.Text(debug_window, wrap=tk.WORD, font=("Consolas", 10))
        scrollbar = tk.Scrollbar(debug_window, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)

        text_widget.insert("1.0", debug_info)
        text_widget.config(state=tk.DISABLED)

        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    except Exception as e:
        messagebox.showerror("调试错误", f"分析Excel文件时出错：\n{str(e)}")


# 创建专门的Excel处理界面
root = tk.Tk()
root.title("Excel植物样方表格整合工具")
root.geometry("600x400")

# 主标题
title_label = tk.Label(
    root,
    text="Excel植物样方表格整合工具",
    font=("微软雅黑", 16, "bold"),
    fg="#2E7D32"
)
title_label.pack(pady=20)

# 说明文本
description = tk.Label(
    root,
    text="专门处理Excel格式的植物样方数据\n自动识别并合并多个独立子表格\n输出按样方编号排序的矩阵",
    font=("微软雅黑", 11),
    fg="#666666",
    justify="center"
)
description.pack(pady=10)

# 处理按钮
process_btn = tk.Button(
    root,
    text="选择Excel文件并处理",
    command=process_excel_file,
    font=("微软雅黑", 12),
    width=20,
    bg="#4CAF50",
    fg="white",
    height=2
)
process_btn.pack(pady=15)

# 调试按钮
debug_btn = tk.Button(
    root,
    text="分析Excel文件结构",
    command=debug_excel_structure,
    font=("微软雅黑", 10),
    width=15,
    bg="#2196F3",
    fg="white"
)
debug_btn.pack(pady=5)

# 使用说明
instructions = tk.Label(
    root,
    text="使用说明:\n"
         "1. Excel文件中应包含多个以'物种'开头的独立表格\n"
         "2. 每个表格应有明确的表头和数据行\n"
         "3. 程序会自动识别所有表格并合并为单一矩阵\n"
         "4. 输出文件将按样方编号(1-1-1, 1-1-2, ...)排序",
    font=("微软雅黑", 9),
    fg="#555555",
    justify="left",
    bg="#F5F5F5",
    padx=10,
    pady=10
)
instructions.pack(pady=20, fill=tk.X, padx=20)

root.mainloop()