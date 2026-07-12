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
import hashlib

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

# 1. 文件二进制MD5（判断完全一模一样的文件）
def get_file_md5(file_path: Path) -> str:
    hash_obj = hashlib.md5()
    with open(file_path, "rb") as f:
        while chunk := f.read(8192):
            hash_obj.update(chunk)
    return hash_obj.hexdigest()

# 2. 文本MD5（判断PDF/OCR文字内容一致，解决元数据不同但内容相同）
def get_text_md5(text: str) -> str:
    return hashlib.md5(text.strip().encode("utf-8")).hexdigest()

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

# 提取单据编号 + 单文件内提单/报关号去重
def extract_all_code(file_path: Path) -> tuple[list[str], list[str], str, str]:
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
        return [], [], "跳过非PDF/图片文件", ""

    # 提取提单号去重
    bl_matches = BL_KEY_REGEX.findall(raw_text)
    temp_bl = [m[1].strip() for m in bl_matches if len(m[1].strip()) >= 6]
    bl_list = list(dict.fromkeys(temp_bl))

    # 提取报关单号去重
    customs_matches = CUSTOMS_REGEX.findall(raw_text)
    temp_cus = [m[1].strip() for m in customs_matches if len(m[1].strip()) >= 16]
    customs_list = list(dict.fromkeys(temp_cus))

    # 文件名兜底提单号
    if not bl_list:
        file_match = FILE_BL_REGEX.search(file_path.name)
        if file_match:
            bl_list = [file_match.group(0).strip()]
            source_type = "文件名兜底匹配提单号"

    text_md5 = get_text_md5(raw_text)
    return bl_list, customs_list, source_type, text_md5

# 文件复制+自定义重命名（兼容提单/报关单两种编号）
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

# 递归多层解压，支持压缩包内子文件夹、嵌套压缩包
def auto_extract_recursive(root: Path, temp_dir: Path):
    temp_dir.mkdir(parents=True, exist_ok=True)
    processed_compress = set()
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

# 全局递归扫描所有层级PDF/图片
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

# 分组归档 + 双重去重：文件MD5 + 文本MD5
def group_and_copy_files(file_list: list[Path], rename_rule: int):
    bl_group = {}
    cus_group = {}
    log_rows = []
    today_str = datetime.now().strftime("%Y-%m-%d")
    seen_file_md5 = set()   # 二进制重复
    seen_text_md5 = set()   # 文本内容重复
    unique_file_list = []

    # 双重去重过滤
    for file in file_list:
        try:
            # 1. 文件二进制完全一致直接跳过
            file_md5 = get_file_md5(file)
            if file_md5 in seen_file_md5:
                print(f"【重复-文件完全一致】跳过：{file.name}")
                continue

            # 2. 读取文本，文本内容一致也跳过（解决元数据不同、肉眼内容相同）
            bls, cuss, src_type, text_md5 = extract_all_code(file)
            if text_md5 in seen_text_md5:
                print(f"【重复-文本内容一致】跳过：{file.name}")
                continue

            # 无重复则加入待处理列表
            seen_file_md5.add(file_md5)
            seen_text_md5.add(text_md5)
            unique_file_list.append(file)
        except Exception as e:
            print(f"读取文件异常，保留文件 {file.name}：{str(e)}")
            unique_file_list.append(file)

    print(f"\n===== 去重统计 =====")
    print(f"原始扫描文件总数：{len(file_list)} 个")
    print(f"去重后待处理单据：{len(unique_file_list)} 个\n")

    # 遍历处理每一份不重复单据
    for file in unique_file_list:
        bl_nums, customs_nums, source, _ = extract_all_code(file)
        print(f"【{source}】{file.name} 提单号：{bl_nums} | 报关单号：{customs_nums}")
        log_rows.append({
            "扫描日期": today_str,
            "原始文件路径": str(file),
            "原始文件名": file.name,
            "识别来源": source,
            "提取提单号": ",".join(bl_nums) if bl_nums else "无",
            "提取报关单号": ",".join(customs_nums) if customs_nums else "无"
        })

        # 区分分类模式：提单号分类 / 报关单号分类
        if rename_rule in [5,6]:
            # 模式5、6：强制以报关单号作为分类主键
            if not customs_nums:
                unknown_dir = Path(OUTPUT_DIR) / "无匹配报关单_人工核对"
                unknown_dir.mkdir(parents=True, exist_ok=True)
                main_cus = ""
                copy_file_rename(file, unknown_dir, rename_rule, "", main_cus)
                continue
            for cus_no in customs_nums:
                safe_cus = clean_dir_name(cus_no)
                target_dir = Path(OUTPUT_DIR) / safe_cus
                target_dir.mkdir(parents=True, exist_ok=True)
                copy_file_rename(file, target_dir, rename_rule, "", cus_no)
                if cus_no not in cus_group:
                    cus_group[cus_no] = []
                cus_group[cus_no].append(file)
        else:
            # 原有模式1/2/3/4：以提单号为主分类
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

    # 导出汇总表
    df = pd.DataFrame(log_rows)
    excel_path = Path(OUTPUT_DIR) / "单据识别汇总清单.xlsx"
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(excel_path, index=False)
    print(f"\n✅ 汇总清单已生成：{excel_path}")

    # 分类统计打印
    print("\n===== 分组统计 =====")
    if rename_rule in [5,6]:
        for cus, files in cus_group.items():
            print(f"报关单号 {cus}：共{len(files)}份单据")
        unknown_count = len([r for r in log_rows if r["提取报关单号"] == "无"])
        print(f"无匹配报关单待人工核对：{unknown_count}份")
    else:
        for bl, files in bl_group.items():
            print(f"提单号 {bl}：共{len(files)}份单据")
        unknown_count = len([r for r in log_rows if r["提取提单号"] == "无"])
        print(f"无匹配提单待人工核对：{unknown_count}份")
    print(f"\n🎉 全部处理完成，输出目录：{OUTPUT_DIR}")

def main():
    root_path = Path(ROOT_DIR)
    out_path = Path(OUTPUT_DIR)
    temp_path = Path(TEMP_UNZIP)
    out_path.mkdir(parents=True, exist_ok=True)

    if not root_path.exists():
        print(f"❌ 错误：单据目录不存在 {ROOT_DIR}")
        return

    # 递归多层解压
    auto_extract_recursive(root_path, temp_path)

    print("===== 单据分组重命名规则选择 =====")
    print("1 - 【提单分类】仅按提单号建文件夹，文件保留原始名称")
    print("2 - 【提单分类】文件重命名：提单号_原文件名")
    print("3 - 【提单分类】文件重命名：报关单号_原文件名")
    print("4 - 【提单分类】文件重命名：当前日期_原文件名")
    print("5 - 【报关单分类】仅按报关单号建文件夹，文件保留原始名称")
    print("6 - 【报关单分类】文件重命名：报关单号_原文件名")
    while True:
        choice = input("请输入数字1/2/3/4/5/6回车：").strip()
        if choice in ["1", "2", "3", "4", "5", "6"]:
            rename_rule = int(choice)
            break
        print("输入错误，请输入1、2、3、4、5、6其中一个数字")

    print(f"\n开始扫描单据目录：{ROOT_DIR}")
    all_files = scan_all_deep_files(root_path, temp_path)
    print(f"原始扫描到文件总数：{len(all_files)} 个")
    group_and_copy_files(all_files, rename_rule)

    # 自动清理临时解压文件夹
    if temp_path.exists():
        shutil.rmtree(temp_path)
        print(f"\n临时解压缓存已自动清理：{TEMP_UNZIP}")

if __name__ == "__main__":
    main()
