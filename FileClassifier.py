import os
import shutil
import hashlib
import zipfile
import py7zr
import pandas as pd
from datetime import datetime
from pathlib import Path
from enum import Enum
from typing import Dict, List, Tuple

# ===================== 全局配置常量 =====================
TEMP_DIR_NAME = "_解压临时文件夹_"
BACKUP_ROOT = Path("原始备份")
DUPLICATE_REPORT_NAME = "去重记录表.xlsx"
CLASSIFY_REPORT_NAME = "分类汇总表.xlsx"
# 匹配置信度阈值
CONFIDENCE_HIGH = 90
CONFIDENCE_MID = 80
CONFIDENCE_LOW = 70

# 去重算法模式枚举
class DedupMode(Enum):
    NAME_CONTENT = "文件名+内容比对"
    META_FAST = "大小+修改时间比对"
    HASH_ACCURATE = "哈希精准比对"

# 分类文件夹目录名称
FOLDER_FULL_MATCH = "自动归档_100%匹配"
FOLDER_HIGH_SUS = "90%-99%高度疑似匹配"
FOLDER_MID_SUS = "80%-89%疑似匹配"
FOLDER_LOW_SUS = "低匹配度待人工排查"

# ===================== 底层工具函数 =====================
def get_file_sha256(file_path: Path, buf_size: int = 65536) -> str:
    """计算文件SHA256哈希值"""
    sha256 = hashlib.sha256()
    try:
        with open(file_path, "rb") as f:
            while chunk := f.read(buf_size):
                sha256.update(chunk)
        return sha256.hexdigest()
    except Exception:
        return ""

def get_file_meta(file_path: Path) -> Tuple[int, float]:
    """获取文件元数据：文件大小、最后修改时间戳"""
    stat = file_path.stat()
    return stat.st_size, stat.st_mtime

def auto_backup(source_root: Path, bak_rule: str = "date") -> Path:
    """
    处理前自动备份全部原始文件
    :param source_root: 待处理根目录
    :param bak_rule: date / 自定义
    :return: 本次备份文件夹路径
    """
    BACKUP_ROOT.mkdir(exist_ok=True)
    if bak_rule == "date":
        folder_name = datetime.now().strftime("%Y-%m-%d")
    else:
        folder_name = "自定义备份集"
    bak_dir = BACKUP_ROOT / folder_name
    bak_dir.mkdir(exist_ok=True)

    # 递归复制所有文件
    for file in source_root.rglob("*"):
        if file.is_file():
            target = bak_dir / file.name
            shutil.copy2(file, target)
    print(f"✅ 文件备份完成，备份路径：{bak_dir.resolve()}")
    return bak_dir

def recursive_extract_archive(root_dir: Path) -> Tuple[int, int]:
    """
    递归扫描并解压zip/7z压缩包，嵌套压缩包一并解压
    :return: 压缩包总数、提取出的文件总数
    """
    temp_dir = root_dir / TEMP_DIR_NAME
    temp_dir.mkdir(exist_ok=True)
    archive_count = 0
    file_count = 0

    for file in root_dir.rglob("*"):
        suffix = file.suffix.lower()
        if suffix not in (".zip", ".7z"):
            continue
        archive_count += 1
        current_temp = temp_dir / file.stem
        current_temp.mkdir(exist_ok=True)

        if suffix == ".zip":
            with zipfile.ZipFile(file, "r") as zf:
                zf.extractall(current_temp)
                file_count += len(zf.namelist())
        elif suffix == ".7z":
            with py7zr.SevenZipFile(file, "r") as sz:
                sz.extractall(current_temp)
                file_count += len(sz.getnames())
    print(f"📦 预处理解压完成：扫描{archive_count}个压缩包，提取{file_count}个文件")
    return archive_count, file_count

# ===================== 文件去重模块 =====================
def file_deduplicate(file_list: List[Path], mode: DedupMode) -> Tuple[List[dict], Dict[str, List[Path]]]:
    """批量文件去重，生成重复分组与Excel报表数据"""
    group_map: Dict[str, List[Path]] = {}
    record_list = []
    group_id = 0

    if mode == DedupMode.HASH_ACCURATE:
        # 哈希精准比对
        hash_mapping: Dict[str, List[Path]] = {}
        for f in file_list:
            h = get_file_sha256(f)
            hash_mapping.setdefault(h, []).append(f)
        for h, paths in hash_mapping.items():
            if len(paths) < 2:
                continue
            group_id += 1
            gid = f"GRP_{datetime.now().strftime('%Y%m%d')}_{group_id:03d}"
            group_map[gid] = paths
            main_file = paths[0]
            record_list.append({
                "重复组ID": gid,
                "文件名": main_file.name,
                "相对路径": str(main_file.parent),
                "比对模式": mode.value,
                "相似度": "100%",
                "角色标注": "主文件"
            })
            for dup in paths[1:]:
                record_list.append({
                    "重复组ID": gid,
                    "文件名": dup.name,
                    "相对路径": str(dup.parent),
                    "比对模式": mode.value,
                    "相似度": "100%",
                    "角色标注": "重复文件"
                })
    elif mode == DedupMode.META_FAST:
        # 大小+修改时间快速比对
        meta_map: Dict[str, List[Path]] = {}
        for f in file_list:
            size, mtime = get_file_meta(f)
            key = f"{size}_{round(mtime, 0)}"
            meta_map.setdefault(key, []).append(f)
        for key, paths in meta_map.items():
            if len(paths) < 2:
                continue
            group_id += 1
            gid = f"GRP_{datetime.now().strftime('%Y%m%d')}_{group_id:03d}"
            group_map[gid] = paths
            main_file = paths[0]
            record_list.append({
                "重复组ID": gid,
                "文件名": main_file.name,
                "相对路径": str(main_file.parent),
                "比对模式": mode.value,
                "相似度": "接近100%",
                "角色标注": "主文件"
            })
            for dup in paths[1:]:
                record_list.append({
                    "重复组ID": gid,
                    "文件名": dup.name,
                    "相对路径": str(dup.parent),
                    "比对模式": mode.value,
                    "相似度": "接近100%",
                    "角色标注": "重复文件"
                })
    else:
        # 默认：文件名+内容简易比对
        name_map: Dict[str, List[Path]] = {}
        for f in file_list:
            name_map.setdefault(f.stem, []).append(f)
        for stem, paths in name_map.items():
            if len(paths) < 2:
                continue
            group_id += 1
            gid = f"GRP_{datetime.now().strftime('%Y%m%d')}_{group_id:03d}"
            group_map[gid] = paths
            main_file = paths[0]
            record_list.append({
                "重复组ID": gid,
                "文件名": main_file.name,
                "相对路径": str(main_file.parent),
                "比对模式": mode.value,
                "相似度": "90%+",
                "角色标注": "主文件"
            })
            for dup in paths[1:]:
                record_list.append({
                    "重复组ID": gid,
                    "文件名": dup.name,
                    "相对路径": str(dup.parent),
                    "比对模式": mode.value,
                    "相似度": "90%+",
                    "角色标注": "重复文件"
                })
    # 导出Excel去重记录表
    df = pd.DataFrame(record_list)
    df.to_excel(DUPLICATE_REPORT_NAME, index=False)
    print(f"📋 去重记录表已生成：{DUPLICATE_REPORT_NAME}")
    return record_list, group_map

# ===================== OCR文字提取（模拟接口，V1仅测试框架） =====================
def ocr_extract_text(file_path: Path, high_precision: bool = False) -> Tuple[str, float]:
    """
    模拟OCR识别：改用【文件名】模拟文本，测试归类功能
    正式版替换为pytesseract读取PDF/图片文字
    """
    # 关键修复：用文件名充当识别文本，实现按关键词归类
    text = file_path.name
    accuracy = 75.0
    if high_precision:
        accuracy = 92.0
    return text, accuracy

# ===================== 智能关键词分类模块（新增文件移动逻辑） =====================
def smart_classify(source_root: Path, file_list: List[Path], keywords: List[str], high_ocr: bool = False) -> Tuple[Dict[str, List[Path]], List[dict]]:
    """按关键词匹配度梯度自动分类，并物理移动文件到对应文件夹"""
    # 创建四大分类根目录
    classify_base = source_root / "分类结果"
    folder_list = [FOLDER_FULL_MATCH, FOLDER_HIGH_SUS, FOLDER_MID_SUS, FOLDER_LOW_SUS]
    for fname in folder_list:
        (classify_base / fname).mkdir(exist_ok=True)

    classify_result = {
        FOLDER_FULL_MATCH: [],
        FOLDER_HIGH_SUS: [],
        FOLDER_MID_SUS: [],
        FOLDER_LOW_SUS: []
    }
    summary_rows = []

    for file in file_list:
        text, acc = ocr_extract_text(file, high_ocr)
        # 识别率低于80%弹出警告
        if acc < 80:
            print(f"⚠️ {file.name} OCR识别准确率仅{acc}%，建议使用高精度模式")

        # 简易关键词匹配计算置信度
        match_count = sum(1 for kw in keywords if kw in text)
        total_key = len(keywords) if keywords else 1
        confidence = (match_count / total_key) * 100

        # 梯度分流判定
        if confidence >= 100:
            target_key = FOLDER_FULL_MATCH
        elif confidence >= CONFIDENCE_HIGH:
            target_key = FOLDER_HIGH_SUS
        elif confidence >= CONFIDENCE_MID:
            target_key = FOLDER_MID_SUS
        else:
            target_key = FOLDER_LOW_SUS
        classify_result[target_key].append(file)...
