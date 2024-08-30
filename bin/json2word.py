#!/usr/bin/env python
# encoding: utf-8

import json
import logging
import os
from datetime import datetime
from docx import Document
from concurrent.futures import ProcessPoolExecutor, as_completed
import gc
import traceback

def process_single_file(json_filename, json_folder):
    try:
        json_path = os.path.join(json_folder, json_filename)

        # 读取 JSON 文件
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # 打开指定的 Word 文档
        doc = Document('/Users/bigyang/myapp/yiheyuan/word/temp.docx')

        # 定义占位符和 JSON 数据字段的映射
        placeholder_mapping = {
            "year": data.get("年"),
            "month": data.get("月"),
            "day": data.get("日"),
            "zongdengjihao": data.get("总登记号"),
            "fenleihao": data.get("分类号"),
            "mingcheng": data.get("名称"),
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
            "beizhu": data.get("备注")
        }

        # 替换段落中的占位符
        for paragraph in doc.paragraphs:
            for placeholder, value in placeholder_mapping.items():
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str(value))

        # 替换表格中的占位符
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for placeholder, value in placeholder_mapping.items():
                            if placeholder in paragraph.text:
                                paragraph.text = paragraph.text.replace(placeholder, str(value))

        # 保存修改后的文档到 ok 文件夹中，名称与 json 文件一致
        output_filename = os.path.join('ok', os.path.splitext(json_filename)[0] + '.docx')
        doc.save(output_filename)
        return json_filename

    except Exception as e:
        # 捕获异常并记录到日志中
        logging.error(f"处理文件 {json_filename} 时出错: {e}\n{traceback.format_exc()}")
        return None

if __name__ == "__main__":
    # 创建 log 和 ok 文件夹
    os.makedirs('log', exist_ok=True)
    os.makedirs('ok', exist_ok=True)

    # 获取当前时间，生成日志文件名称
    log_filename = datetime.now().strftime("log/%Y-%m-%d_%H-%M-%S.log")

    # 配置日志，将日志写入文件
    logging.basicConfig(
        filename=log_filename,
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

    # 获取用户指定的 JSON 文件夹路径（仅输入一次）
    json_folder = input("请输入 JSON 文件夹路径: ")

    # 获取所有 JSON 文件列表
    json_files = [f for f in os.listdir(json_folder) if f.endswith('.json')]
    total_files = len(json_files)
    batch_size = 100  # 每次处理的文件数量
    generated_files_count = 0

    # 使用进程池进行并行处理
    with ProcessPoolExecutor(max_workers=4) as executor:  # 可以调整 max_workers 来控制并发进程数
        futures = [executor.submit(process_single_file, json_filename, json_folder) for json_filename in json_files]
        for i, future in enumerate(as_completed(futures)):
            try:
                result = future.result()
                if result:
                    generated_files_count += 1
            except Exception as e:
                logging.error(f"处理文件时出现异常: {e}\n{traceback.format_exc()}")

            # 显示进度
            print(f"处理进度: {i+1}/{total_files}", end="\r")

            # 定期释放内存
            if (i + 1) % batch_size == 0:
                gc.collect()

    # 提示用户程序已完成
    print(f"\n程序运行完毕，一共生成 {generated_files_count} 个文件，请查看。")

