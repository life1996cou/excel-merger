import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

def merge_excel():
    try:
        # 获取用户输入
        skip_header = int(entry_header.get())
        skip_footer = int(entry_footer.get())
        folder_path = entry_folder.get()
        width_mode = width_choice.get()  # 获取列宽模式选择

        # 检查文件夹路径
        if not os.path.isdir(folder_path):
            messagebox.showerror("错误", "文件夹路径无效！")
            return

        # 遍历文件夹中的 Excel 文件并合并
        all_data = []
        column_widths = {}  # 记录每列最大宽度

        for file in os.listdir(folder_path):
            if file.endswith((".xlsx", ".xls")):
                file_path = os.path.join(folder_path, file)
                
                # 读取 Excel 文件
                df = pd.read_excel(file_path, engine='openpyxl', header=None)
                # 跳过表头和表尾
                df = df.iloc[skip_header:-skip_footer if skip_footer > 0 else None]
                
                # 处理标题行
                if not all_data:
                    header = df.iloc[0]
                    df = df[1:]
                else:
                    df = df[1:]
                
                all_data.append(df)

                # 仅在模式1时记录列宽
                if width_mode == 1:
                    wb = load_workbook(file_path)
                    ws = wb.active
                    for idx, col in enumerate(ws.columns, 1):
                        max_length = max(len(str(cell.value)) for cell in col)
                        column_letter = get_column_letter(idx)
                        current_width = column_widths.get(column_letter, 0)
                        column_widths[column_letter] = max(max_length, current_width)

        # 合并数据
        if all_data:
            merged_df = pd.concat(all_data, ignore_index=True)
            merged_df.columns = header  # 设置标题
            
            # 保存合并后的文件
            output_path = os.path.join(folder_path, "合并.xlsx")
            merged_df.to_excel(output_path, index=False, engine='openpyxl')
            
            # 调整格式
            wb = load_workbook(output_path)
            ws = wb.active
            
            # 设置列宽
            if width_mode == 1:  # 原表格列宽
                for col_letter, width in column_widths.items():
                    ws.column_dimensions[col_letter].width = width + 2
            else:  # 自动调整列宽
                for idx, col in enumerate(merged_df.columns, 1):
                    max_length = max(
                        merged_df[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    ws.column_dimensions[get_column_letter(idx)].width = max_length + 2
            
            # 设置居中格式
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            
            wb.save(output_path)
            messagebox.showinfo("成功", f"文件已保存为：{output_path}")
        else:
            messagebox.showerror("错误", "文件夹内没有 Excel 文件！")
    except Exception as e:
        messagebox.showerror("错误", f"发生错误：{str(e)}")

# 创建图形界面
root = tk.Tk()
root.title("Excel 合并工具")

# 第一行：跳过表头
ttk.Label(root, text="跳过表头行数:").grid(row=0, column=0, padx=10, pady=5)
entry_header = ttk.Entry(root)
entry_header.grid(row=0, column=1, padx=10, pady=5)
entry_header.insert(0, "0")

# 第二行：跳过表尾
ttk.Label(root, text="跳过末尾行数:").grid(row=1, column=0, padx=10, pady=5)
entry_footer = ttk.Entry(root)
entry_footer.grid(row=1, column=1, padx=10, pady=5)
entry_footer.insert(0, "0")

# 第三行：选择文件夹
ttk.Label(root, text="选择文件夹:").grid(row=2, column=0, padx=10, pady=5)
entry_folder = ttk.Entry(root, width=40)
entry_folder.grid(row=2, column=1, padx=10, pady=5)

def select_folder():
    folder = filedialog.askdirectory()
    if folder:
        entry_folder.delete(0, tk.END)
        entry_folder.insert(0, folder)

ttk.Button(root, text="浏览...", command=select_folder).grid(row=2, column=2, padx=10, pady=5)

# 第四行：列宽模式选择
width_choice = tk.IntVar(value=2)  # 默认选择自动调整
ttk.Label(root, text="列宽模式:").grid(row=3, column=0, padx=10, pady=5)
ttk.Radiobutton(root, text="原表格列宽（不稳定）", variable=width_choice, value=1).grid(row=3, column=1, sticky="w")
ttk.Radiobutton(root, text="自动调整列宽（推荐）", variable=width_choice, value=2).grid(row=4, column=1, sticky="w")

# 合并按钮
ttk.Button(root, text="开始合并", command=merge_excel).grid(row=5, column=1, pady=10)

root.mainloop()