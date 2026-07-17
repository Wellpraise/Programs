#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
发货单对账工具 v1.2 - 完整打包版
====================================

功能：
  1. 扫描箱子照片中的 FBA 条码
  2. 从文件夹名推断 SKU
  3. 与第一步汇总表联动对账
  4. 串号检测 + 清晰度评估

作者：WorkBuddy
日期：2026-07-17

依赖安装：
  pip install pyzbar opencv-python pandas openpyxl pillow

注意事项：
  - pyzbar 需要 ZBar 库支持
  - Windows 需要 Visual C++ Redistributable
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime


# ==================== 配置区域 ====================
FIRST_STAGE_OUTPUT_FOLDER = r"E:\WORKBUDDY\验货\A0151\汇总结果"
DEFAULT_OUTPUT_FOLDER = r"E:\WORKBUDDY\验货\A0151\对账结果"
CLARITY_THRESHOLD = 15.0  # 清晰度阈值（Laplacian方差）
# =================================================


def collect_image_files(folder):
    """递归收集所有图片文件"""
    exts = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.webp')
    files = []
    if os.path.isdir(folder):
        for root, dirs, fnames in os.walk(folder):
            for fname in fnames:
                if fname.lower().endswith(exts) and 'thumbs.db' not in fname.lower():
                    full_path = os.path.join(root, fname)
                    rel_path = os.path.relpath(full_path, folder)
                    files.append((full_path, rel_path))
    return sorted(files, key=lambda x: x[1])


def parse_photo_path(rel_path):
    """
    从路径推断 SKU
    示例: AE/162件/SOB001/IMG_xxx.jpg → SKU = SOB001
    """
    parts = rel_path.replace('\\', '/').split('/')
    folders = parts[:-1]
    sku = ''
    folder_hierarchy = '/'.join(folders)

    if folders:
        for word in folders[-1].split():
            if word.startswith(('SOK', 'SOB')):
                sku = word
                break

    return {'sku': sku, 'folder_hierarchy': folder_hierarchy, 'all_folders': folders}


def scan_barcodes(image_path):
    """
    扫描照片中的条码
    返回: [条码内容列表]
    """
    from pyzbar.pyzbar import decode
    from PIL import Image

    barcodes = []
    try:
        img = Image.open(image_path)

        # 原图扫描
        results = decode(img)
        for r in results:
            data = r.data.decode('utf-8').strip()
            if data and (data.upper().startswith('FBA') or len(data) >= 8):
                barcodes.append(data)

        # 尝试缩放后扫（有时缩小反而更容易识别模糊条码）
        scaled = img.resize((img.width // 2, img.height // 2))
        results2 = decode(scaled)
        for r in results2:
            data = r.data.decode('utf-8').strip()
            if data and (data.upper().startswith('FBA') or len(data) >= 8) and data not in barcodes:
                barcodes.append(data)

    except Exception:
        pass

    return barcodes


def classify_barcodes(barcodes):
    """
    分类扫描到的条码
    返回: {fba_codes, labels, skus, other}
    """
    result = {'fba_codes': [], 'labels': [], 'skus': [], 'other': []}

    for code in barcodes:
        code_upper = code.upper()

        # FBA 编码（如 FBA19HBZ8XRK 或 FBA19HBZ8XRKU000012）
        if 'FBA' in code_upper and len(code) >= 11:
            result['fba_codes'].append(code)
        # 亚马逊标签码（10位字母数字，如 X003Y64QXZ）
        elif len(code) == 10 and code.isalnum() and any(c.isalpha() for c in code):
            result['labels'].append(code)
        # SKU（SOK/SOB 开头）
        elif code.startswith(('SOK', 'SOB')):
            result['skus'].append(code)
        else:
            result['other'].append(code)

    return result


def check_image_clarity(image_path):
    """检测图片清晰度"""
    try:
        from PIL import Image
        import cv2
        import numpy as np

        img_pil = Image.open(image_path)
        img_np = np.array(img_pil)

        if len(img_np.shape) == 3:
            gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
        else:
            gray = img_np

        score = cv2.Laplacian(gray, cv2.CV_64F).var()
        return score, score >= CLARITY_THRESHOLD

    except ImportError:
        return 0, False


def load_first_stage_data():
    """加载第一步汇总表"""
    import pandas as pd

    folders_to_search = [FIRST_STAGE_OUTPUT_FOLDER]
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for parent in [script_dir, os.path.dirname(script_dir), os.path.dirname(os.path.dirname(script_dir))]:
        result_folder = os.path.join(parent, "汇总结果")
        if os.path.isdir(result_folder):
            folders_to_search.append(result_folder)

    for folder in folders_to_search:
        if not os.path.isdir(folder):
            continue
        for fname in os.listdir(folder):
            if fname.startswith("发货单_SKU标签汇总") and fname.endswith('.xlsx'):
                fpath = os.path.join(folder, fname)
                try:
                    df = pd.read_excel(fpath, engine='openpyxl')
                    if '箱子编码' in df.columns and len(df) > 0:
                        return df
                except:
                    pass

    return None


def run_matching(image_folder, summary_df, output_folder):
    """执行条码扫描 + 对账匹配"""
    import pandas as pd

    image_files = collect_image_files(image_folder)
    if not image_files:
        return None, "该目录中没有找到图片文件", [], [], []

    all_records = []
    errors = []
    status_log = []

    status_log.append(f"找到 {len(image_files)} 张图片\n")
    status_log.append("=" * 90 + "\n")

    for i, (full_path, rel_path) in enumerate(image_files, 1):
        fname = os.path.basename(full_path)
        try:
            # 1. 路径解析
            parsed = parse_photo_path(rel_path)
            path_sku = parsed['sku']

            # 2. 条码扫描
            barcodes = scan_barcodes(full_path)
            classified = classify_barcodes(barcodes)

            # 3. 清晰度检测
            clarity_score, is_clear = check_image_clarity(full_path)

            # 4. 从汇总表中查 FBA
            fba_code = ''
            match_status = '待确认'
            serial_status = '未知'

            if summary_df is not None and not summary_df.empty:
                # 优先用扫描到的 FBA 编码匹配
                if classified['fba_codes']:
                    fba_code = classified['fba_codes'][0]
                    # FBA编码格式: FBA19HBB4QFKU000005 -> BS-FBA19HBB4QFKU000005
                    # 尝试完整匹配 + 带前缀匹配
                    matched = summary_df[summary_df['箱子编码'] == fba_code]
                    if len(matched) == 0:
                        matched = summary_df[summary_df['箱子编码'].str.startswith(fba_code[:13], na=False)]
                    if len(matched) > 0:
                        fba_code = matched.iloc[0]['箱子编码']
                    else:
                        fba_code = ''
                # 或用路径 SKU 匹配
                elif path_sku:
                    matched = summary_df[summary_df['SKU号'] == path_sku]
                    if len(matched) > 0:
                        fba_code = matched.iloc[0]['箱子编码']

                # 串号判断
                if fba_code:
                    all_skus_for_fba = summary_df[summary_df['箱子编码'] == fba_code]['SKU号'].unique()
                    if len(all_skus_for_fba) > 1:
                        serial_status = '串号'
                    else:
                        serial_status = '正常无串号'

            # 5. 判定匹配状态
            if fba_code and is_clear:
                match_status = '✅ 匹配成功'
            elif fba_code and not is_clear:
                match_status = '⚠️ 有FBA但照片模糊'
            elif classified['fba_codes']:
                match_status = '⚠️ 扫描到FBA但未匹配到汇总表'
            else:
                match_status = '❌ 未识别到FBA条码'

            # 6. 记录
            record = {
                '原图文件名': fname,
                '相对路径': rel_path,
                '路径推断SKU': path_sku,
                '扫描FBA编码': fba_code if fba_code else (classified['fba_codes'][0] if classified['fba_codes'] else ''),
                '所有扫描条码': '; '.join(barcodes),
                '扫描条码总数': len(barcodes),
                '扫描到的标签码': '; '.join(classified['labels']),
                '串号状态': serial_status,
                '匹配状态': match_status,
                '图片清晰度': f'{clarity_score:.2f}',
                '是否清晰': '是' if is_clear else '否',
                '备注': path_sku or '路径中无SKU信息'
            }
            all_records.append(record)

            status_log.append(
                f"[{i}/{len(image_files)}] {fname}\n"
                f"    SKU:{path_sku or '无'} | FBA:{fba_code[:20] or classified['fba_codes'][0][:20] if classified['fba_codes'] else '无'} | "
                f"条码:{len(barcodes)}个 | 清晰度:{clarity_score:.1f} | 匹配:{match_status}\n"
            )

        except Exception as e:
            status_log.append(f"[错误] {fname} - {str(e)}\n")
            errors.append((fname, str(e)))

    # 生成报表
    if all_records:
        result_df = pd.DataFrame(all_records)
        os.makedirs(output_folder, exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        total_filename = f"对账总表_{timestamp}.xlsx"
        total_path = os.path.join(output_folder, total_filename)
        result_df.to_excel(total_path, index=False, engine='openpyxl')

        for status_text, file_prefix in [('✅ 匹配成功', '正常匹配'), ('⚠️ 有FBA但照片模糊', '模糊照片'),
                                          ('❌ 未识别到FBA条码', '无条码'), ('⚠️ 扫描到FBA但未匹配到汇总表', '未匹配FBA')]:
            cat_df = result_df[result_df['匹配状态'] == status_text]
            if len(cat_df) > 0:
                cat_filename = f"{file_prefix}_{timestamp}.xlsx"
                cat_df.to_excel(os.path.join(output_folder, cat_filename), index=False, engine='openpyxl')

        serial_df = result_df[result_df['串号状态'] == '串号']
        if len(serial_df) > 0:
            serial_df.to_excel(os.path.join(output_folder, f"串号_{timestamp}.xlsx"), index=False, engine='openpyxl')

        return all_records, status_log, errors, [total_path]

    return [], [], [], []


def run_gui():
    """GUI 界面"""
    root = tk.Tk()
    root.title("发货单对账工具 v1.2 - 条码扫描版")
    root.geometry("700x500")
    root.resizable(True, True)

    style = ttk.Style()
    style.configure('Title.TLabel', font=('Microsoft YaHei UI', 14, 'bold'))
    style.configure('Info.TLabel', font=('Microsoft YaHei UI', 9))

    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(main_frame, text="📷 发货单对账工具 v1.2", style='Title.TLabel').pack(pady=(0, 10))
    ttk.Label(main_frame, text="扫描箱子照片中的 FBA 条码 + 路径推断 SKU\n与第一步汇总表联动对账校验",
              justify=tk.CENTER, style='Info.TLabel').pack(pady=(0, 15))

    progress_var = tk.DoubleVar()
    ttk.Progressbar(main_frame, variable=progress_var, maximum=100).pack(fill=tk.X, pady=10)

    status_label = ttk.Label(main_frame, text="就绪", style='Title.TLabel')
    status_label.pack(pady=(0, 10))

    log_text = tk.Text(main_frame, height=10, font=('Consolas', 8))
    log_text.pack(fill=tk.BOTH, expand=True, pady=10)

    btn_frame = ttk.Frame(main_frame)
    btn_frame.pack(fill=tk.X, pady=10)

    def log_msg(msg):
        log_text.config(state=tk.NORMAL)
        log_text.insert(tk.END, msg + '\n')
        log_text.see(tk.END)
        log_text.config(state=tk.DISABLED)
        root.update()

    def on_run():
        image_folder = filedialog.askdirectory(title="选择箱子照片文件夹")
        if not image_folder:
            return

        output_folder = filedialog.askdirectory(title="选择输出文件夹", initialdir=DEFAULT_OUTPUT_FOLDER)
        if not output_folder:
            return

        status_label.config(text="正在加载汇总表...")
        log_msg("========================================")
        log_msg(f"照片: {image_folder}")

        summary_df = load_first_stage_data()
        if summary_df is None or summary_df.empty:
            log_msg("⚠️ 未找到汇总表")
            summary_df = None
        else:
            log_msg(f"✅ 汇总表: {len(summary_df)} 条")

        progress_var.set(0)
        status_label.config(text="正在扫描条码...")

        try:
            all_records, status_log, errs, paths = run_matching(image_folder, summary_df, output_folder)

            if all_records:
                import pandas as pd
                records_df = pd.DataFrame(all_records)
                ok = len(records_df[records_df['匹配状态'] == '✅ 匹配成功'])
                blur = len(records_df[records_df['匹配状态'] == '⚠️ 有FBA但照片模糊'])
                nomatch = len(records_df[records_df['匹配状态'] == '❌ 未识别到FBA条码'])
                serial = len(records_df[records_df['串号状态'] == '串号'])

                msg = f"完成！\n\n总图片: {len(all_records)}\n匹配成功: {ok}\n模糊照片: {blur}\n无条码: {nomatch}\n串号: {serial}\n失败: {len(errs)}"
                log_msg("=" * 60)
                log_msg(msg)
                log_msg("=" * 60)
                messagebox.showinfo("结果", msg)
                try:
                    os.startfile(output_folder)
                except:
                    pass
        except Exception as e:
            import traceback
            log_msg(f"❌ {str(e)}\n{traceback.format_exc()}")
            messagebox.showerror("错误", str(e))

    ttk.Button(btn_frame, text="▶ 开始对账", command=on_run).pack(side=tk.LEFT, padx=5, expand=True)

    def on_help():
        messagebox.showinfo("使用说明",
            "📷 扫码识别方案\n\n"
            "1. 依赖: pip install pyzbar opencv-python pandas openpyxl pillow\n"
            "2. 扫描箱子照片中的 FBA 条码\n"
            "3. 从文件夹名推断 SKU（如 SOB001）\n"
            "4. 与汇总表比对匹配\n"
            "5. 串号检测 + 清晰度评估\n\n"
            "输出: 总表 + 正常/模糊/无条码/串号 分类文件")

    ttk.Button(btn_frame, text="📖 帮助", command=on_help).pack(side=tk.LEFT, padx=5, expand=True)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
