import tkinter as tk
from tkinter import filedialog, END
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
import os
import json
import re
import asyncio
from concurrent.futures import ProcessPoolExecutor
import aiofiles

class ExcelToJsonConverter(ttk.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title("提取 Excel 中的数据为 JSON 文件")
        self.geometry("700x500")
        self.create_widgets()

    def create_widgets(self):
        # 设置全局字体
        default_font = ("Microsoft YaHei", 12)  # 使用微软雅黑作为默认字体
        self.option_add("*Font", default_font)

        # 创建自定义按钮样式
        self.custom_style = ttk.Style()
        self.custom_style.configure("TButton", font=default_font, foreground="black", background="light gray")
        self.custom_style.map("TButton", background=[("active", "gray")])

        # 文件选择框
        self.file_frame = ttk.Frame(self)
        self.file_frame.pack(pady=20, padx=20, fill=X)

        self.file_entry = ttk.Entry(self.file_frame, width=50, font=default_font)
        self.file_entry.pack(side=LEFT, expand=YES, fill=X)

        self.file_button = ttk.Button(self.file_frame, text="选择Excel文件", command=self.select_file)
        self.file_button.pack(side=RIGHT, padx=(10, 0))

        # 输出目录选择框
        self.output_frame = ttk.Frame(self)
        self.output_frame.pack(pady=(0, 20), padx=20, fill=X)

        self.output_entry = ttk.Entry(self.output_frame, width=50, font=default_font)
        self.output_entry.pack(side=LEFT, expand=YES, fill=X)

        self.output_button = ttk.Button(self.output_frame, text="选择输出目录", command=self.select_output_dir)
        self.output_button.pack(side=RIGHT, padx=(10, 0))

        # 转换按钮
        self.convert_button = ttk.Button(self, text="开始转换", command=self.start_conversion, style="success.TButton")
        self.convert_button.pack(pady=10)

        # 进度条和进度标签
        self.progress_frame = ttk.Frame(self)
        self.progress_frame.pack(pady=20, padx=20, fill=X)

        self.progress = ttk.Progressbar(self.progress_frame, length=500, mode='determinate', style="success.Horizontal.TProgressbar")
        self.progress.pack(side=LEFT, expand=YES, fill=X)

        self.progress_label = ttk.Label(self.progress_frame, text="0 / 0", font=("Microsoft YaHei", 10))
        self.progress_label.pack(side=RIGHT, padx=(10, 0))

        # 状态标签
        self.status_label = ttk.Label(self, text="准备就绪", font=("Microsoft YaHei", 12))
        self.status_label.pack(pady=10)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="选择Excel文件"
        )
        if file_path:
            self.file_entry.delete(0, END)
            self.file_entry.insert(0, file_path)

    def select_output_dir(self):
        output_dir = filedialog.askdirectory(title="选择输出目录")
        if output_dir:
            self.output_entry.delete(0, END)
            self.output_entry.insert(0, output_dir)

    def start_conversion(self):
        file_path = self.file_entry.get()
        output_dir = self.output_entry.get()

        if not file_path or not output_dir:
            self.status_label.config(text="请选择Excel文件和输出目录")
            return

        self.convert_button.config(state="disabled")
        self.status_label.config(text="正在转换...")
        self.progress['value'] = 0
        self.progress_label.config(text="0 / 0")

        asyncio.run(self.convert_excel_to_json(file_path, output_dir))

    async def convert_excel_to_json(self, file_path, output_dir):
        try:
            data = self.read_excel(file_path)
            total_records = len(data)
            batch_size = 500

            os.makedirs(output_dir, exist_ok=True)

            processed_records = 0
            with ProcessPoolExecutor(max_workers=8) as executor:
                loop = asyncio.get_event_loop()
                
                for i in range(0, total_records, batch_size):
                    chunk = data.iloc[i:i + batch_size]
                    tasks = []
                    
                    for index, row in chunk.iterrows():
                        file_name, row_dict = await loop.run_in_executor(executor, self.process_row, index, row)
                        tasks.append(self.write_json(output_dir, file_name, row_dict))
                        
                        processed_records += 1
                        self.progress['value'] = (processed_records / total_records) * 100
                        self.progress_label.config(text=f"{processed_records} / {total_records}")
                        self.update_idletasks()
                    
                    await asyncio.gather(*tasks)

            self.status_label.config(text=f"转换完成！数据已导出到 {output_dir} 目录。")
        except Exception as e:
            self.status_label.config(text=f"转换过程中出错：{str(e)}")
        finally:
            self.convert_button.config(state="normal")

    @staticmethod
    def read_excel(file_path):
        _, ext = os.path.splitext(file_path)
        if ext.lower() in ['.xls', '.xlsx']:
            return pd.read_excel(file_path, dtype=str)
        else:
            raise ValueError("不支持的文件格式，请使用 .xls 或 .xlsx 文件。")

    @staticmethod
    def clean_filename(filename):
        filename = re.sub(r'[\u200B-\u200D\uFEFF]', '', filename)
        filename = re.sub(r'[^\w\s-]', '', filename)
        return filename.strip()

    @staticmethod
    def process_row(index, row):
        row_dict = row.to_dict()
        file_name = row_dict.get('总登记号', f'row_{index+1}')
        file_name = ExcelToJsonConverter.clean_filename(file_name)
        for key in row_dict:
            row_dict[key] = str(row_dict[key])
        return file_name, row_dict

    @staticmethod
    async def write_json(output_dir, file_name, row_dict):
        json_file_path = os.path.join(os.path.abspath(output_dir), f'{file_name}.json')
        async with aiofiles.open(json_file_path, 'w', encoding='utf-8') as json_file:
            await json_file.write(json.dumps(row_dict, ensure_ascii=False, indent=4))

if __name__ == "__main__":
    app = ExcelToJsonConverter()
    app.mainloop()
