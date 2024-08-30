#!/usr/bin/env python
# encoding: utf-8

import pandas as pd
import os
import json
import re
import aiofiles
import asyncio
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm

# 加载Excel文件
file_path = '/Users/bigyang/myapp/yiheyuan/excel/source.xls'  # 替换为实际的Excel文件路径

# 创建保存JSON文件的目录（如果不存在则创建）
output_dir = '/Users/bigyang/myapp/yiheyuan/json/'
os.makedirs(output_dir, exist_ok=True)

# 函数：清理文件名中的非法字符和不可见字符
def clean_filename(filename):
    # 移除不可见字符（如零宽度空格、控制字符等）
    filename = re.sub(r'[\u200B-\u200D\uFEFF]', '', filename)  # 移除零宽度字符
    filename = re.sub(r'[^\w\s-]', '', filename)  # 移除非字母、数字、下划线、连字符和空格的字符
    filename = filename.strip()  # 去除首尾空格
    return filename

# 异步函数：将字典写入JSON文件
async def write_json(file_name, row_dict):
    json_file_path = os.path.join(output_dir, f'{file_name}.json')
    async with aiofiles.open(json_file_path, 'w', encoding='utf-8') as json_file:
        await json_file.write(json.dumps(row_dict, ensure_ascii=False, indent=4))

# 多线程处理函数
def process_row(index, row):
    row_dict = row.to_dict()
    file_name = row_dict.get('总登记号', f'row_{index+1}')  # 如果“总登记号”为空，使用行号作为文件名
    file_name = clean_filename(file_name)
    return file_name, row_dict

# 主函数：处理Excel数据
async def main():
    # 读取整个Excel文件
    data = pd.read_excel(file_path)

    # 计算总行数
    total_records = len(data)

    # 设置批次大小
    batch_size = 500  # 可根据系统内存进行调整

    with ThreadPoolExecutor(max_workers=4) as executor:  # 根据CPU核心数量调整线程数
        loop = asyncio.get_event_loop()
        with tqdm(total=total_records) as pbar:  # 初始化进度条
            for i in range(0, total_records, batch_size):
                chunk = data.iloc[i:i + batch_size]  # 手动分批
                tasks = []
                for index, row in chunk.iterrows():
                    file_name, row_dict = await loop.run_in_executor(executor, process_row, index, row)
                    tasks.append(write_json(file_name, row_dict))
                    pbar.update(1)  # 每处理一行，更新进度条
                await asyncio.gather(*tasks)

if __name__ == "__main__":
    asyncio.run(main())
    print(f"数据已导出到 {output_dir} 目录。")
