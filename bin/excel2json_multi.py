#!/usr/bin/env python
# encoding: utf-8

import pandas as pd
import os
import json
import re
import aiofiles
import asyncio
from concurrent.futures import ProcessPoolExecutor  # 使用多进程
from tqdm import tqdm

# Excel 文件路径（支持 .xls 和 .xlsx 文件）
file_path = '/Users/bigyang/myapp/yiheyuan/excel/source.xlsx'  # 替换为实际的 Excel 文件路径

# JSON 文件保存目录（如果不存在则创建）
output_dir = '/Users/bigyang/myapp/yiheyuan/json/'
os.makedirs(output_dir, exist_ok=True)

# 函数：清理文件名中的非法字符和不可见字符
def clean_filename(filename):
    # 移除不可见字符（如零宽度空格、控制字符等）
    filename = re.sub(r'[\u200B-\u200D\uFEFF]', '', filename)  # 移除零宽度字符
    filename = re.sub(r'[^\w\s-]', '', filename)  # 移除非字母、数字、下划线、连字符和空格的字符
    filename = filename.strip()  # 去除首尾空格
    return filename

# 异步函数：将字典写入 JSON 文件
async def write_json(file_name, row_dict):
    # 构建 JSON 文件的完整路径
    json_file_path = os.path.join(output_dir, f'{file_name}.json')
    
    # 使用异步方式打开文件并写入 JSON 数据
    async with aiofiles.open(json_file_path, 'w', encoding='utf-8') as json_file:
        await json_file.write(json.dumps(row_dict, ensure_ascii=False, indent=4))

# 多进程处理函数：处理每一行的 Excel 数据，转换为字典并生成文件名
def process_row(index, row):
    row_dict = row.to_dict()
    
    # 提取“总登记号”作为文件名，如果为空则使用行号作为文件名
    file_name = row_dict.get('总登记号', f'row_{index+1}')
    
    # 清理文件名，确保合法性
    file_name = clean_filename(file_name)
    
    # 将所有列数据转换为字符串类型，以确保写入 JSON 时为文本格式
    for key in row_dict:
        row_dict[key] = str(row_dict[key])
    
    return file_name, row_dict

# 函数：判断文件扩展名并读取 Excel 文件
def read_excel(file_path):
    _, ext = os.path.splitext(file_path)
    
    # 根据文件扩展名选择正确的读取方式
    if ext == '.xls':
        return pd.read_excel(file_path, dtype=str)  # 读取 .xls 文件，确保数据为字符串
    elif ext == '.xlsx':
        return pd.read_excel(file_path, dtype=str)  # 读取 .xlsx 文件，确保数据为字符串
    else:
        raise ValueError("不支持的文件格式，请使用 .xls 或 .xlsx 文件。")

# 主函数：处理 Excel 数据
async def main():
    # 读取 Excel 文件，并将所有数据加载为 DataFrame
    try:
        data = read_excel(file_path)  # 调用读取函数，支持 .xls 和 .xlsx
    except Exception as e:
        print(f"读取 Excel 文件时出错：{str(e)}")
        return

    # 计算总行数
    total_records = len(data)

    # 设置批次大小，避免占用过多内存
    batch_size = 500  # 根据系统内存情况可调整

    # 使用多进程池来处理数据，max_workers 可以设置为系统的 CPU 核心数
    with ProcessPoolExecutor(max_workers=8) as executor:  # 适合 8 核 CPU
        loop = asyncio.get_event_loop()
        with tqdm(total=total_records) as pbar:  # 初始化进度条，显示处理进度
            for i in range(0, total_records, batch_size):
                chunk = data.iloc[i:i + batch_size]  # 手动分批读取数据
                tasks = []
                
                # 遍历当前批次的每一行
                for index, row in chunk.iterrows():
                    # 使用多进程池并行处理每一行数据
                    file_name, row_dict = await loop.run_in_executor(executor, process_row, index, row)
                    
                    # 异步写入每一行数据到单独的 JSON 文件
                    tasks.append(write_json(file_name, row_dict))
                    
                    # 每处理一行，进度条更新一次
                    pbar.update(1)
                
                # 等待所有异步任务完成
                await asyncio.gather(*tasks)

if __name__ == "__main__":
    # 运行主程序
    asyncio.run(main())
    print(f"数据已导出到 {output_dir} 目录。")

