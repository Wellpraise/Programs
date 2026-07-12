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
import rarfile

filterwarnings("ignore")

# ==================== 固定路径，无改动 ====================
ROOT_DIR = r"/Users/inbest/Desktop/test"
OUTPUT_DIR = r"/Users/inbest/Desktop/提单分组结果"
TEMP_UNZIP = r"/Users/inbest/Desktop/temp_unzip_cache"
# ==========================================================

# 提单、报关正则
BL_KEY_REGEX = re.compile(
    r'(B/L\s*No[:：]\s*|BL\s*No[:：]\s*|提单号[:：]\s*|提运单号[:：]\s*)([A-Z0-9\-]{6,30})',
    re.IGNORECASE
)
CUSTOMS_REGEX = re.compile(
    r'(报关单号[:：]\s*|预录入编号[:：]\s*)([0-9]{16,20})',
    re.IGNORECASE
)
FILE_BL_REGEX = re.compile(r'[A-Z0-9]{6,}[-A-Z0-9]+')

# 清理文件名非法字符
def clean_dir_name(name: str) -> str:
    illegal_chars = r'\/:*?"<>|'
    for c in illegal_chars:
        name = name.replace(c, "_")
    return name.strip()

# PDF原生文字提取
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

# PDF图片OCR识别
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

# 图片OCR识别
def ocr_img_text(img_path: Path) -> str:
    try:
        img = Image.open(img_path)
        text = pytesseract.image_to_string(img, lang="chi_sim+eng", config="--psm 6")
        return re.sub(r'\s+', ' ', text)
    except Exception:
        return ""

# 提取单据编号 + 自动去重
def extract_all_code(file_path: Path) -> tuple[list[str], list[str], str]:
    suffix = file_path.suffix.lower()
    source_type = "PDF正文识别"
    bl_list = []
    customs_list = []
    raw_text = ""

    if suffix == ".pdf":
        raw_text = read_pdf_text(file_path)
        if not raw_text.strip():
            raw_text = ocr_pdf_text(file_path)
            source_type = "PDF图片OCR识别"
    elif suffix in (".jpg", ".jpeg", ".png"):
        raw_text = ocr_img_text(file_path)
        source_type = "图片OCR识别"
    else:
        return [], [], "跳过非PDF/图片文件"

    # 提取提单号并自动去重
    bl_matches = BL_KEY_REGEX.findall(raw_text)
    temp_bl = [m[1].strip() for m in bl_matches if len(m[1].strip()) >= 6]
    bl_list = list(dict.fromkeys(temp_bl))

    # 提取报关单号并自动去重
    customs_matches = CUSTOMS_REGEX.findall(raw_text)
    temp_cus = [m[1].strip() for m in customs_matches if len(m[1].strip()) >= 16]
    customs_list = list(dict.fromkeys(temp_cus))

    # 正文无提单号，从文件名兜底匹配
    if not bl_list:
        file_match = FILE_BL_REGEX.search(file_path.name)
        if file_match:
            bl_list = [file_match.group(0).strip()]
            source_type = "文件名兜底匹配提单号"

    return bl_list, customs_list, source_type

# 文件复制+自定义重命名
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

    dst_file = dst_dir / new_name
    idx = 1
    while dst_file.exists():
        dst_file = dst_dir / f"{dst_file.stem}_{idx}{suffix}"
        idx += 1
    shutil.copy2(src, dst_file)

# 【增强：递归多层解压，支持压缩包内子文件夹、嵌套压缩包】
def auto_extract_recursive(root: Path, temp_dir: Path):
    temp_dir.mkdir(parents=True, exist_ok=True)
    processed_compress = set()
    # 循环多次处理嵌套压缩包，直到无新压缩包
    while True:
        all_compress = list(root.rglob("*.zip")) + list(root.rglob("*.rar")) + list(temp_dir.rglob("*.zip")) + list(temp_dir.rglob("*.rar"))
        new_compress = []
        for comp_file in all_compress:
            if str(comp_file) in processed_compress:
                continue
            new_compress.append(comp_file)
        if not new_compress:
            break
        for comp_file in new_compress:
            processed_compress.add(str(comp_file))
            try:
                sub_temp = temp_dir / clean_dir_name(f"{comp_file.stem}_{len(processed_compress)}")
                sub_temp.mkdir(exist_ok=True)
                if comp_file.suffix.lower() == ".zip":
                    with zipfile.ZipFile(comp_file, 'r') as zf:
                        zf.extractall(sub_temp)
                elif comp_file.suffix.lower() == ".rar":
                    with rarfile.RarFile(comp_file, 'r') as rf:
                        rf.extractall(sub_temp)
                print(f"已解压：{comp_file.name}")
            except Exception as e:
                print(f"解压失败 {comp_file.name}：{str(e)}")

# 【增强：全局递归扫描所有层级PDF/图片，不漏子文件夹】
def scan_all_deep_files(root: Path, temp_dir: Path) -> list[Path]:
    support_suffix = {".pdf", ".jpg", ".jpeg", ".png"}
    file_list = []
    # 源目录所有层级文件
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in support_suffix:
            file_list.append(path)
    # 临时解压目录所有层级文件
    for path in temp_dir.rglob("*"):
        if path.is_file() and path.suffix.lower() in support_suffix:
            file_list.append(path)
    return file_list

# 分组归档逻辑（完全沿用旧版）
def group_and_copy_files(file_list: list[Path], rename_rule: int):
    bl_group = {}
    log_rows = []
    today_str = datetime.now().strftime("%Y-%m-%d")

    for file in file_list:
        bl_nums, customs_nums, source = extract_all_code(file)
        print(f"【{source}】{file.name} 提单号：{bl_nums} | 报关单号：{customs_nums}")
        log_rows.append({
            "扫描日期": today_str,
            "原始文件路径": str(file),
            "原始文件名": file.name,
            "识别来源": source,
            "提取提单号": ",".join(bl_nums) if bl_nums else "无",
            "提取报关单号": ",".join(customs_nums) if customs_nums else "无"
        })

        if not bl_nums:
            unknown_dir = Path(OUTPUT_DIR) / "无匹配提单_人工核对"
            unknown_dir.mkdir(parents=True, exist_ok=True)
            main_cus = customs_nums[0] if customs_nums else ""
            copy_file_rename(file, unknown_dir, rename_rule, "", main_cus)
            continue

        for bl in bl_nums:
            safe_bl = clean_dir_name(bl)
            target_dir = Path(OUTPUT_DIR) / safe_bl
            target_dir.mkdir(parents=True, exist_ok=True)
            main_cus = customs_nums[0] if customs_nums else ""
            copy_file_rename(file, target_dir, rename_rule, bl, main_cus)
            if bl not in bl_group:
                bl_group[bl] = []
            bl_group[bl].append(file)

    df = pd.DataFrame(log_rows)
    excel_path = Path(OUTPUT_DIR) / "单据识别汇总清单.xlsx"
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(excel_path, index=False)
    print(f"\n✅ 汇总清单已生成：{excel_path}")

    print("\n===== 提单分组统计 =====")
    for bl, files in bl_group.items():
        print(f"提单号 {bl}：共{len(files)}份单据")
    unknown_count = len([r for r in log_rows if r["提取提单号"] == "无"])
    print(f"待人工核对文件：{unknown_count}份")
    print(f"\n🎉 全部处理完成，输出目录：{OUTPUT_DIR}")

def main():
    root_path = Path(ROOT_DIR)
    out_path = Path(OUTPUT_DIR)
    temp_path = Path(TEMP_UNZIP)
    out_path.mkdir(parents=True, exist_ok=True)

    if not root_path.exists():
        print(f"❌ 错误：单据目录不存在 {ROOT_DIR}")
        return

    # 递归多层解压（修复嵌套压缩包、子文件夹）
    auto_extract_recursive(root_path, temp_path)

    print("===== 单据分组重命名规则选择 =====")
    print("1 - 仅按提单号建文件夹，文件保留原始名称")
    print("2 - 文件重命名：提单号_原文件名（你需要的模式）")
    print("3 - 文件重命名：报关单号_原文件名")
    print("4 - 文件重命名：当前日期_原文件名")
    while True:
        choice = input("请输入数字1/2/3/4回车：").strip()
        if choice in ["1", "2", "3", "4"]:
            rename_rule = int(choice)
            break
        print("输入错误，请输入1、2、3、4其中一个数字")

    print(f"\n开始扫描单据目录：{ROOT_DIR}")
    all_files = scan_all_deep_files(root_path, temp_path)
    print(f"共扫描到待处理单据：{len(all_files)} 个\n")
    group_and_copy_files(all_files, rename_rule)

    # 自动清理临时解压文件夹
    if temp_path.exists():
        shutil.rmtree(temp_path)
        print(f"\n临时解压缓存已自动清理：{TEMP_UNZIP}")

if __name__ == "__main__":
    main()
