import os
import json
import logging
from docx import Document
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime
from rich.progress import Progress, TextColumn, BarColumn, TimeRemainingColumn
from multiprocessing import cpu_count
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox

class WordGeneratorApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title("Word Generator")
        self.geometry("600x400")
        self.create_widgets()

    def create_widgets(self):
        # Frame for input fields
        input_frame = ttk.Frame(self, padding=10)
        input_frame.pack(fill=BOTH, expand=YES)

        # Template file selection
        ttk.Label(input_frame, text="Word 模板文件:").grid(row=0, column=0, sticky=W, pady=5)
        self.template_entry = ttk.Entry(input_frame, width=50)
        self.template_entry.grid(row=0, column=1, pady=5)
        ttk.Button(input_frame, text="浏览", command=self.select_template).grid(row=0, column=2, padx=5)

        # JSON folder selection
        ttk.Label(input_frame, text="JSON 文件夹:").grid(row=1, column=0, sticky=W, pady=5)
        self.json_entry = ttk.Entry(input_frame, width=50)
        self.json_entry.grid(row=1, column=1, pady=5)
        ttk.Button(input_frame, text="浏览", command=self.select_json_folder).grid(row=1, column=2, padx=5)

        # Output folder selection
        ttk.Label(input_frame, text="输出文件夹:").grid(row=2, column=0, sticky=W, pady=5)
        self.output_entry = ttk.Entry(input_frame, width=50)
        self.output_entry.grid(row=2, column=1, pady=5)
        ttk.Button(input_frame, text="浏览", command=self.select_output_folder).grid(row=2, column=2, padx=5)

        # Start button
        self.start_button = ttk.Button(self, text="开始处理", command=self.start_processing, bootstyle=SUCCESS)
        self.start_button.pack(pady=10)

        # Progress bar
        self.progress = ttk.Progressbar(self, length=500, mode='determinate', style='info.Horizontal.TProgressbar')
        self.progress.pack(pady=10)

        # Status label
        self.status_label = ttk.Label(self, text="就绪")
        self.status_label.pack(pady=5)

    def select_template(self):
        filename = filedialog.askopenfilename(filetypes=[("Word Document", "*.docx")])
        if filename:
            self.template_entry.delete(0, END)
            self.template_entry.insert(0, filename)

    def select_json_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.json_entry.delete(0, END)
            self.json_entry.insert(0, folder)

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_entry.delete(0, END)
            self.output_entry.insert(0, folder)

    def start_processing(self):
        template_path = self.template_entry.get()
        json_folder = self.json_entry.get()
        output_folder = self.output_entry.get()

        if not all([template_path, json_folder, output_folder]):
            messagebox.showerror("错误", "请填写所有必要的路径")
            return

        self.start_button.config(state=DISABLED)
        self.status_label.config(text="处理中...")
        self.progress['value'] = 0

        # 在新线程中运行处理过程
        self.after(100, lambda: self.run_processing(template_path, json_folder, output_folder))

    def run_processing(self, template_path, json_folder, output_folder):
        try:
            # 设置日志
            log_dir = os.path.join(output_folder, 'log')
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            log_file = os.path.join(log_dir, f'{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.log')
            logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(message)s')

            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            json_files = self.get_sorted_json_files(json_folder)
            
            if not json_files:
                messagebox.showerror("错误", "没有找到任何 JSON 文件")
                self.reset_ui()
                return

            total_files = len(json_files)
            processed_files = 0

            with ProcessPoolExecutor(max_workers=cpu_count()) as executor:
                futures = [executor.submit(self.process_single_file, json_file, template_path, output_folder) for json_file in json_files]
                
                for future in as_completed(futures):
                    future.result()
                    processed_files += 1
                    self.update_progress(processed_files, total_files)

            messagebox.showinfo("完成", f"处理完成，共处理 {total_files} 个文件")
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中发生错误: {str(e)}")
        finally:
            self.reset_ui()

    def update_progress(self, processed, total):
        progress_value = int((processed / total) * 100)
        self.progress['value'] = progress_value
        self.status_label.config(text=f"已处理: {processed}/{total}")
        self.update_idletasks()

    def reset_ui(self):
        self.start_button.config(state=NORMAL)
        self.status_label.config(text="就绪")

    @staticmethod
    def get_sorted_json_files(json_folder):
        json_files = [os.path.join(json_folder, f) for f in os.listdir(json_folder) if f.endswith('.json')]
        return sorted(json_files, key=lambda x: os.path.basename(x))

    @staticmethod
    def map_json_to_placeholders(data):
        return {
            "year": data.get("年"),
            "month": data.get("月"),
            "day": data.get("日"),
            "zongdengjihao": data.get("总登记号"),
            "fenleihao": data.get("分类号"),
            "name": data.get("名称"),
            "niandai": data.get("年代"),
            "jianshu": data.get("件数"),
            "danwei": data.get("单位"),
            "chicun": data.get("尺寸"),
            "zhongliang": data.get("重量"),
            "zhidi": data.get("质地"),
            "wancanqingkuang": data.get("完残情况"),
            "laiyuan": data.get("来源"),
            "ruguanpingzhenghao": data.get("入馆凭证号"),
            "zhuxiaopingzhenghao": data.get("注销凭证号"),
            "jibie": data.get("级别"),
            "beizhu": data.get("备注"),
            "fuzeren": data.get("负责人"),
            "danganbianhao": data.get("档案编号"),
            "xingzhuangneirongmiaoshu": data.get("形状内容描述"),
            "dangqianbaocuntiaojian": data.get("当前保存条件"),
            "mingjitiba": data.get("铭记题跋")
        }

    @staticmethod
    def process_single_file(json_file, template_path, output_folder):
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            doc = Document(template_path)
            placeholders = WordGeneratorApp.map_json_to_placeholders(data)

            for p in doc.paragraphs:
                for key, value in placeholders.items():
                    if value:
                        p.text = p.text.replace(f'{key}', str(value))

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for key, value in placeholders.items():
                            if value:
                                cell.text = cell.text.replace(f'{key}', str(value))

            output_file = os.path.join(output_folder, os.path.basename(json_file).replace('.json', '.docx'))
            doc.save(output_file)

            logging.info(f'成功生成文件: {output_file}')
            return True
        except Exception as e:
            logging.error(f"处理文件 {json_file} 时出错: {str(e)}")
            return False

if __name__ == "__main__":
    app = WordGeneratorApp()
    app.mainloop()
