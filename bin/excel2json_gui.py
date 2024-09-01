#!/usr/bin/env python
# encoding: utf-8

import wx
import pandas as pd
import os
import json
import re
import aiofiles
import asyncio
from concurrent.futures import ThreadPoolExecutor
import logging
from datetime import datetime

# Configure logging
log_dir = '/home/bigyang/python_bigyang/yiheyuan/log/'
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f'excel2json_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
logging.basicConfig(filename=log_file, level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to clean filenames
def clean_filename(filename):
    filename = re.sub(r'[\u200B-\u200D\uFEFF]', '', filename)
    filename = re.sub(r'[^\w\s-]', '', filename)
    filename = filename.strip()
    return filename

# Function to clean strings
def clean_string(s):
    if isinstance(s, str):
        s = re.sub(r'[\u200B-\u200D\uFEFF]', '', s)  # Remove invisible characters
        s = re.sub(r'^\s+|\s+$', '', s)  # Remove leading and trailing whitespace
    return s

# Async function to write JSON
async def write_json(file_name, row_dict, output_dir):
    json_file_path = os.path.join(output_dir, f'{file_name}.json')
    async with aiofiles.open(json_file_path, 'w', encoding='utf-8') as json_file:
        await json_file.write(json.dumps(row_dict, ensure_ascii=False, indent=4))

# Function to process each row
def process_row(index, row):
    try:
        row_dict = {key: clean_string(value) for key, value in row.to_dict().items()}
        file_name = row_dict.get('总登记号', f'row_{index+1}')
        file_name = clean_filename(file_name)
        return file_name, row_dict
    except Exception as e:
        logging.error(f"Error processing row {index}: {e}")
        return None, None

# Main function to process Excel data
async def process_excel(file_path, output_dir, progress_callback):
    batch_size = 100
    total_records = 0
    file_extension = os.path.splitext(file_path)[1].lower()

    try:
        if file_extension == '.xlsx':
            data = pd.read_excel(file_path, engine='openpyxl', dtype=str)
            total_records = len(data)
            headers = data.columns.tolist()

            def read_chunk(start, end):
                return data.iloc[start:end]

        elif file_extension == '.xls':
            data = pd.read_excel(file_path, engine='xlrd', dtype=str)
            total_records = len(data)
            headers = data.columns.tolist()

            def read_chunk(start, end):
                return data.iloc[start:end]

        else:
            raise ValueError("不支持的文件格式！请选择 xls 或者 xlsx 文件！")

        with ThreadPoolExecutor(max_workers=4) as executor:
            loop = asyncio.get_event_loop()
            tasks = []
            for chunk_start in range(0, total_records, batch_size):
                chunk_end = min(chunk_start + batch_size, total_records)
                chunk = read_chunk(chunk_start, chunk_end)

                for index, row in chunk.iterrows():
                    file_name, row_dict = await loop.run_in_executor(executor, process_row, index, row)
                    if file_name is not None:
                        tasks.append(write_json(file_name, row_dict, output_dir))
                        progress_callback(chunk_start, total_records)

            await asyncio.gather(*tasks)

    except Exception as e:
        logging.error(f"Error processing Excel file: {e}")
        raise

# GUI part
class MyApp(wx.App):
    def OnInit(self):
        self.frame = MyFrame(None, title="Excel 数据导出到独立 JSON 文件")
        self.frame.Show()
        return True

class MyFrame(wx.Frame):
    def __init__(self, *args, **kw):
        super(MyFrame, self).__init__(*args, **kw, size=(800, 600))  # Increase window size

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        self.label_excel = wx.StaticText(panel, label="请选择 Excel 文件位置：")
        vbox.Add(self.label_excel, proportion=0, flag=wx.EXPAND | wx.ALL, border=10)
        self.file_picker = wx.FilePickerCtrl(panel, message="选择 Excel 文件位置")
        vbox.Add(self.file_picker, proportion=0, flag=wx.EXPAND | wx.ALL, border=10)

        self.label_json = wx.StaticText(panel, label="请选择 JSON 文件存放位置：")
        vbox.Add(self.label_json, proportion=0, flag=wx.EXPAND | wx.ALL, border=10)
        self.output_dir_picker = wx.DirPickerCtrl(panel, message="选择 JSON 文件的保存目录")
        vbox.Add(self.output_dir_picker, proportion=0, flag=wx.EXPAND | wx.ALL, border=10)

        self.start_button = wx.Button(panel, label="开始处理")
        vbox.Add(self.start_button, proportion=0, flag=wx.CENTER | wx.ALL, border=10)
        self.start_button.Bind(wx.EVT_BUTTON, self.on_start)

        self.progress_bar = wx.Gauge(panel, range=100, size=(500, 30))  # Increase progress bar size
        vbox.Add(self.progress_bar, proportion=0, flag=wx.EXPAND | wx.ALL, border=10)

        panel.SetSizer(vbox)
        panel.Layout()

        # Center the window
        self.Centre()

        self.Bind(wx.EVT_CLOSE, self.on_close)

    def update_progress(self, current, total):
        percentage = int((current / total) * 100)
        self.progress_bar.SetValue(percentage)

    async def run_processing(self):
        file_path = self.file_picker.GetPath()
        output_dir = self.output_dir_picker.GetPath()

        if file_path and output_dir:
            try:
                await process_excel(file_path, output_dir, self.update_progress)
                wx.CallAfter(self.on_processing_complete)
            except Exception as e:
                logging.error(f"Error occurred during processing: {str(e)}")
                wx.CallAfter(self.on_processing_complete)
        else:
            wx.MessageBox("请选择 Excel 文件位置和 JSON 文件存放目录。", "错误", wx.OK | wx.ICON_ERROR)

    def on_start(self, event):
        asyncio.run(self.run_processing())

    def on_processing_complete(self):
        wx.MessageBox("处理完成！", "信息", wx.OK | wx.ICON_INFORMATION)
        self.Close(True)

    def on_close(self, event):
        self.Destroy()

if __name__ == "__main__":
    app = MyApp()
    app.MainLoop()
