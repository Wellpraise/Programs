import os
import shutil
import re
import pandas as pd
from pathlib import Path
from warnings import filterwarnings
import pdfplumber
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
from datetime import datetime
import zipfile
import hashlib
import subprocess

filterwarnings("ignore")

# ==================== Windows配置路径 ====================
ROOT_DIR = Path(r"E:\Coprograms\input")
OUTPUT_DIR = Path(r"E:\Coprograms\output")
TEMP_UNZIP = Path(r"E:\Coprograms\temp_unzip")
# 这里直接使用你 D 盘的 7-Zip 路径
SEVEN_ZIP_PATH = r"D:\7z\7z.exe" 
# ==========================================================

# 匹配规则
BL_KEY_REGEX = re.compile(
    r'(B/L\s*No[:：\s*|BL\s*No[:：\s*|提单号:：\s*|提单号[:：\s*)([A-Z0-9\-]{6,30})',
    re.IGNORECASE
)
CUSTOMS_REGEX = re.compile(
    r'(海关单号[:：\s*|预录入单号:：\s*)([0-9]{16,20})',
    re.IGNORECASE
)
FILE_BL_REGEX = re.compile(r'[A-Z0-9]{6,}[-A-Z0-9]+')

def clean_dir_name(name: str) -> str:
    illegal_chars = r'\/:*?"<>|'
    for c in illegal_chars:
        name = name.replace(c, "_")
    return name.strip()

def get_file_md5(file_path: Path) -> str:
    hash_obj = hashlib.md5()
    with open(file_path, "rb") as f:
        while chunk := f.read(8192):
            hash_obj.update(chunk)
    return hash_obj.hexdigest()

def get_text_md5(text: str) -> str:
    return hashlib.md5(text.strip().encode("utf-8"))

def read_pdf_text(pdf_path: Path) -> str:
    text_all = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text_all += page_text + " "
        return re.sub(r'\s+', ' ', text_all)
    except Exception:
        return ""

def ocr_pdf_text(pdf_path: Path) -> str:
    text_all = ""
    try:
        pages = convert_from_path(pdf_path)
        for img in pages:
            text = pytesseract.image_to_string(img, lang="chi_sim+eng", config="--psm 6")
            text_all += re.sub(r'\s+', ' ', text) + " "
        return text_all
    except Exception:
        return ""

def ocr_img_text(img_path: Path) -> str:
    try:
        img = Image.open(img_path)
        text = pytesseract.image_to_string(img, lang="chi_sim+eng", config="--psm 6")
        return re.sub(r'\s+', ' ', text)
    except Exception:
        return ""

def extract_all_code(file_path: Path) -> tuple[list[str], list[str], str, str]:
    suffix = file_path.suffix.lower()
    source_type = "PDF文本识别"
    bl_list = []
    customs_list = []
    raw_text = ""

    if suffix == ".pdf":
        raw_text = read_pdf_text(file_path)
        if not raw_text.strip():
            raw_text = ocr_pdf_text(file_path)
            source_type = "PDF图像OCR识别"
    elif suffix in (".jpg", ".jpeg", ".png"):
        raw_text = ocr_img_text(file_path)
        source_type = "图像OCR识别"
    else:
        return [], [], "非PDF/图像文件", ""

    bl_matches = BL_KEY_REGEX.findall(raw_text)
    temp_bl = [m[1].strip() for m in bl_matches if len(m[1].strip()) >= 6]
    bl_list = list(dict.fromkeys(temp_bl))

    customs_matches = CUSTOMS_REGEX.findall(raw_text)
    temp_cus = [m[1].strip() for m in customs_matches if len(m[1].strip()) >= 16]
    customs_list = list(dict.fromkeys(temp_cus))

    if not bl_list:
        file_match = FILE_BL_REGEX.search(file_path.name)
        if file_match:
            bl_list = [file_match.group(0).strip()]
            source_type = "文件名包含提单号"

    text_md5 = get_text_md5(raw_text)
    return bl_list, customs_list, source_type, text_md5

def copy_file_rename(src: Path, dst_dir: Path, rule: int, bl_no: str, customs_no: str):
    stem = src.stem
    suffix = src.suffix
    new_name = src.name

    if rule == 2 and bl_no:
        safe_bl = clean_dir_name(bl_no)
        new_name = f"{safe_bl}_{stem}{suffix}"
    elif rule == 3 and customs_no:
        safe_cus = clean_dir_name(customs_no)
        new_name = f"{safe_cus}_{stem}{suffix}"
    elif rule == 4:
        today = datetime.now().strftime("%Y%m%d")
        new_name = f"{today}_{stem}{suffix}"
    elif rule == 6 and customs_no:
        safe_cus = clean_dir_name(customs_no)
        new_name = f"{safe_cus}_{stem}{suffix}"

    dst_file = dst_dir / new_name
    idx = 1
    while dst_file.exists():
        dst_file = dst_dir / f"{dst_file.stem}_{idx}{suffix}"
        idx += 1
    shutil.copy2(src, dst_file)

def auto_extract_recursive(root: Path, temp_dir: Path):
    temp_dir.mkdir(parents=True, exist_ok=True)
    processed_compress = set()
    
    # 支持的压缩格式
    compress_extensions = {".zip", ".rar", ".7z"}
    
    while True:
        all_files = list(root.rglob("*")) + list(temp_dir.rglob("*"))
        found_compress = [f for f in all_files if f.suffix.lower() in compress_extensions and str(f) not in processed_compress]
        
        if not found_compress:
            break
            
        for comp_file in found_compress:
            processed_compress.add(str(comp_file))
            try:
                sub_temp = temp_dir / clean_dir_name(f"{comp_file.stem}_{len(processed_compress)}")
                sub_temp.mkdir(exist_ok=True)
                
                # 使用 7-Zip 进行解压 (x 表示解压并保留目录结构)
                cmd = [SEVEN_ZIP_PATH, "x", str(comp_file), f"-o{sub_temp}", "-y"]
                subprocess.run(cmd, capture_output=True, check=True)
                print(f"已解压: {comp_file.name}")
            except Exception as e:
                print(f"解压失败 {comp_file.name}, {str(e)}")

def scan_all_deep_files(root: Path, temp_dir: Path) -> list[Path]:
    support_suffix = {".pdf", ".jpg", ".jpeg", ".png"}
    file_list = []
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in support_suffix:
            file_list.append(path)
    for path in temp_dir.rglob("*"):
        if path.is_file() and path.suffix.lower() in support_suffix:
            file_list.append(path)
    return file_list

def group_and_copy_files(file_list: list[Path], rename_rule: int):
    bl_group = {}
    cus_group = {}
    log_rows = []
    today_str = datetime.now().strftime("%Y-%m-%d")
    seen_file_md5 = set()
    seen_text_md5 = set()
    unique_file_list = []

    for file in file_list:
        try:
            file_md5 = get_file_md5(file)
            if file_md5 in seen_file_md5:
                print(f"【重复文件完全一致】跳过: {file.name}")
                continue
            bls, cuss, src_type, text_md5 = extract_all_code(file)
            if text_md5 in seen_text_md5:
                print(f"【重复文本内容一致】跳过: {file.name}")
                continue
            seen_file_md5.add(file_md5)
            seen_text_md5.add(text_md5)
            unique_file_list.append(file)
        except Exception as e:
            print(f"读取文件异常，保留文件 {file.name}, {str(e)}")
            unique_file_list.append(file)

    print(f"\n===== 去重统计 =====")
    print(f"原始扫描文件总数: {len(file_list)} 个")
    print(f"去重后实际处理单号: {len(unique_file_list)} 个\n")

    for file in unique_file_list:
        bl_nums, customs_nums, source, _ = extract_all_code(file)
        print(f"[{source}] {file.name} 提单号: {bl_nums} | 海关单号: {customs_nums}")
        log_rows.append({
            "扫描日期": today_str,
            "原始文件路径": str(file),
            "原始文件名": file.name,
            "识别来源": source,
            "提取提单号": ",".join(bl_nums) if bl_nums else "无",
            "提取海关单号": ",".join(customs_nums) if customs_nums else "无"
        })

        if rename_rule in [5,6]:
            if not customs_nums:
                unknown_dir = OUTPUT_DIR / "无海关单_人工核对"
                unknown_dir.mkdir(parents=True, exist_ok=True)
                copy_file_rename(file, unknown_dir, rename_rule, "", "")
                continue
            for cus_no in customs_nums:
                safe_cus = clean_dir_name(cus_no)
                target_dir = OUTPUT_DIR / safe_cus
                target_dir.mkdir(parents=True, exist_ok=True)
                copy_file_rename(file, target_dir, rename_rule, "", cus_no)
                if cus_no not in cus_group:
                    cus_group[cus_no] = []
                cus_group[cus_no].append(file)
        else:
            if not bl_nums:
                unknown_dir = OUTPUT_DIR / "无提单号_人工核对"
                unknown_dir.mkdir(parents=True, exist_ok=True)
                main_cus = customs_nums[0] if customs_nums else ""
                copy_file_rename(file, unknown_dir, rename_rule, "", main_cus)
                continue
            for bl in bl_nums:
                safe_bl = clean_dir_name(bl)
                target_dir = OUTPUT_DIR / safe_bl
                target_dir.mkdir(parents=True, exist_ok=True)
                main_cus = customs_nums[0] if customs_nums else ""
                copy_file_rename(file, target_dir, rename_rule, bl, main_cus)
                if bl not in bl_group:
                    bl_group[bl] = []
                bl_group[bl].append(file)

    df = pd.DataFrame(log_rows)
    excel_path = OUTPUT_DIR / "单号识别汇总清单.xlsx"
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(excel_path, index=False)
    print(f"\n✅ 汇总清单已生成: {excel_path}")

    print("\n===== 分类结果 =====")
    if rename_rule in [5,6]:
        for cus, files in cus_group.items():
            print(f"海关单号 {cus}: 共{len(files)}份文件")
        unknown_count = len([r for r in log_rows if r["提取海关单号"] == "无"])
        print(f"无海关单号文件共计: {unknown_count}份")
    else:
        for bl, files in bl_group.items():
            print(f"提单号 {bl}: 共{len(files)}份文件")
        unknown_count = len([r for r in log_rows if r["提取提单号"] == "无"])
        print(f"无提单号文件共计: {unknown_count}份")
    print(f"\n🎉 全部处理完成！输出结果在: {OUTPUT_DIR}")

def main():
    root_path = ROOT_DIR
    out_path = OUTPUT_DIR
    temp_path = TEMP_UNZIP
    out_path.mkdir(parents=True, exist_ok=True)

    if not root_path.exists():
        print(f"❌ 错误: 数据源目录不存在 {ROOT_DIR}")
        return

    auto_extract_recursive(root_path, temp_path)

    print("===== 单号分类重命名规则选择 =====")
    print("1 - 【提单分类】仅按提单号建文件夹，文件名保持原样")
    print("2 - 【提单分类】文件名重命名: 提单号_原文件名")
    print("3 - 【提单分类】文件名重命名: 海关号_原文件名")
    print("4 - 【提单分类】文件名重命名: 当前日期_原文件名")
    print("5 - 【海关分类】仅按海关号建文件夹，文件名保持原样")
    print("6 - 【海关分类】文件名重命名: 海关号_原文件名")
    while True:
        choice = input("请输入数字 1/2/3/4/5/6 回车: ").strip()
        if choice in ["1", "2", "3", "4", "5", "6"]:
            rename_rule = int(choice)
            break
        print("输入错误，请输入1-6之间的数字")

    print(f"\n开始扫描目录: {ROOT_DIR}")
    all_files = scan_all_deep_files(root_path, temp_path)
    print(f"原始扫描到文件总数: {len(all_files)} 个")
    group_and_copy_files(all_files, rename_rule)

    if temp_path.exists():
        shutil.rmtree(temp_path)
        print(f"\n临时解压缓存已自动清理: {TEMP_UNZIP}")

if __name__ == "__main__":
    main()