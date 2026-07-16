#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel 发货单 SKU 汇总脚本 v3
功能：按照 VBA 逻辑提取发货通知单数据
- 读取每个 .xls 文件中的 '发货单' 工作表
- A1 = 箱子编码(FBA编号), A4 = 发货单号
- 从第13行开始扫描 H列(SKU号) 和 J列(对应标签)
- J列标签采用"向上覆盖"机制：读到新标签就刷新，直到下一行新标签
- 过滤：SKU不为空、不含中文字符、有有效标签才导出
- 输出4列：箱子编码、发货单号、SKU号、对应标签
"""

import os
import re
import unicodedata
import xlrd
import pandas as pd

# ==================== 配置区域 ====================
INPUT_FOLDER = r"E:\验货\A0151\input the list"       # 输入文件夹
OUTPUT_FOLDER = r"E:\验货\A0151\汇总结果_v3"          # 输出文件夹
OUTPUT_FILENAME = "发货单_SKU标签汇总.xlsx"            # 输出文件名
DATA_START_ROW = 12  # 第13行，索引从0开始所以是12
LABEL_OVERRIDE_THRESHOLD = 10  # 连续空白行数达到此值，清空标签
# =================================================


def has_chinese(text):
    """判断字符串是否包含中文字符（Unicode > 255）"""
    if not text:
        return False
    for char in text:
        if ord(char) > 255:
            return True
    return False


def get_cell_value(ws, row, col):
    """安全获取单元格值（兼容合并单元格和不同数据类型）"""
    try:
        cell = ws.cell(row, col)
        if cell.ctype == xlrd.XL_CELL_EMPTY:
            return ""
        elif cell.ctype == xlrd.XL_CELL_TEXT:
            return str(cell.value).strip()
        elif cell.ctype == xlrd.XL_CELL_NUMBER:
            # 如果是整数则去掉小数点
            val = cell.value
            if val == int(val):
                return str(int(val))
            return str(val)
        else:
            return str(cell.value).strip()
    except:
        return ""


def main():
    # 检查输入目录
    if not os.path.isdir(INPUT_FOLDER):
        print(f"错误：输入目录不存在 -> {INPUT_FOLDER}")
        return

    # 获取所有 .xls 和 .xlsx 文件
    all_files = os.listdir(INPUT_FOLDER)
    xls_files = [f for f in all_files if f.lower().endswith(('.xls', '.xlsx'))]

    if not xls_files:
        print("错误：输入目录中没有找到 .xls 或 .xlsx 文件")
        return

    print(f"找到 {len(xls_files)} 个文件")
    print("=" * 90)

    # 存储所有数据
    all_records = []
    errors = []
    skipped_no_sheet = []

    for filename in sorted(xls_files):
        filepath = os.path.join(INPUT_FOLDER, filename)
        try:
            wb = xlrd.open_workbook(filepath)
            sheet_names = wb.sheet_names()

            # 检查是否有 '发货单' 工作表
            if '发货单' not in sheet_names:
                skipped_no_sheet.append(filename)
                continue

            ws = wb.sheet_by_name('发货单')

            # ========== 提取头部信息 ==========
            box_code = get_cell_value(ws, 0, 0)  # A1 - 箱子编码/FBA编号
            ship_no = get_cell_value(ws, 3, 0)   # A4 - 发货单号

            if not box_code and not ship_no:
                print(f"[跳过] {filename} - 头部信息为空")
                continue

            # ========== 扫描数据行 ==========
            active_label = ""  # 当前有效的标签
            empty_row_count = 0  # 连续空白行数计数器
            file_record_count = 0

            for i_row in range(DATA_START_ROW, min(1000, ws.nrows)):
                sku = get_cell_value(ws, i_row, 7)   # H列
                raw_label = get_cell_value(ws, i_row, 9)  # J列

                # 如果读到新标签，刷新 active_label
                if raw_label:
                    active_label = raw_label
                    empty_row_count = 0  # 有新数据，重置计数器

                # 如果 H列(SKU) 为空，累加空白计数器
                if not sku:
                    empty_row_count += 1
                    # 连续10行无SKU，认为明细结束，清空标签
                    if empty_row_count >= LABEL_OVERRIDE_THRESHOLD:
                        active_label = ""
                else:
                    empty_row_count = 0  # 有SKU，重置空白计数器

                # ========== 写入条件 ==========
                # 1. SKU 不为空
                # 2. SKU 不含中文字符
                # 3. 有有效标签
                if sku and not has_chinese(sku) and active_label:
                    all_records.append({
                        '箱子编码': box_code,
                        '发货单号': ship_no,
                        'SKU号': sku,
                        '对应标签': active_label,
                        '源文件名': filename
                    })
                    file_record_count += 1

            print(f"[成功] {filename} - 提取 {file_record_count} 条 | FBA: {box_code[:20]}... | 单号: {ship_no[:30]}...")

        except Exception as e:
            error_msg = f"{filename} - 读取失败: {str(e)}"
            print(f"[错误] {error_msg}")
            errors.append((filename, str(e)))

    # ========== 打印总结 ==========
    print("=" * 90)
    print(f"总记录数: {len(all_records)}")
    print(f"成功文件: {len(xls_files) - len(errors) - len(skipped_no_sheet)}")
    print(f"跳过文件(无发货单表): {len(skipped_no_sheet)}")
    print(f"失败文件: {len(errors)}")

    if skipped_no_sheet:
        print(f"\n跳过的文件:")
        for f in skipped_no_sheet:
            print(f"  - {f}")

    if errors:
        print(f"\n失败文件详情:")
        for fname, err in errors:
            print(f"  - {fname}: {err}")

    # ========== 生成汇总表 ==========
    if all_records:
        result_df = pd.DataFrame(all_records)

        # 创建输出目录
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        # 保存结果
        output_path = os.path.join(OUTPUT_FOLDER, OUTPUT_FILENAME)
        result_df.to_excel(output_path, index=False, engine='openpyxl')

        print(f"\n✅ 汇总结果已保存到: {output_path}")
        print(f"\n📋 输出列:")
        for i, col in enumerate(result_df.columns, 1):
            print(f"   {i}. {col}")
        print(f"\n📊 数据预览（前10行）:")
        print(result_df.head(10).to_string(index=False))
        
        # 统计每个文件的记录数
        print(f"\n📦 各文件记录数:")
        file_counts = result_df.groupby('源文件名').size().sort_values(ascending=False)
        for fname, count in file_counts.items():
            print(f"   {fname}: {count} 条")
    else:
        print("\n⚠️  没有提取到任何有效数据")


if __name__ == "__main__":
    main()
