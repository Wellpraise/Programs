import os
import sys
import argparse
from openpyxl import Workbook
from datetime import datetime


def get_all_folders(root_dir):
    """递归获取指定目录下的所有文件夹（含子文件夹）"""
    all_dirs = []

    for dirpath, dirnames, filenames in os.walk(root_dir):
        for dirname in dirnames:
            full_path = os.path.join(dirpath, dirname)
            rel_path = os.path.relpath(full_path, root_dir)
            depth = rel_path.count(os.sep) + 1
            all_dirs.append({
                'depth': depth,
                'rel_path': rel_path,
                'name': dirname,
                'full_path': full_path
            })

    # 按层级排序
    all_dirs.sort(key=lambda x: (x['depth'], x['rel_path']))
    return all_dirs


def create_excel(all_dirs, output_path):
    """生成Excel文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "文件夹列表"

    # 表头
    ws.append(["序号", "层级", "相对路径", "文件夹名称", "完整路径"])

    # 表头加粗
    for cell in ws[1]:
        from openpyxl.styles import Font
        cell.font = Font(bold=True)

    # 写入数据
    for i, d in enumerate(all_dirs, 2):
        ws.append([i - 1, d['depth'], d['rel_path'], d['name'], d['full_path']])

    # 列宽
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 80
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 120

    wb.save(output_path)


def ensure_output_dir(path):
    """确保输出文件的父目录存在"""
    dir = os.path.dirname(path)
    if dir and not os.path.exists(dir):
        os.makedirs(dir, exist_ok=True)
    return path


def main():
    parser = argparse.ArgumentParser(
        description="文件夹递归提取工具 - 将指定目录下的所有子文件夹名导出为Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  python 文件夹提取工具.py -s "E:\\验货\\A0151"
  python 文件夹提取工具.py -s "E:\\验货\\A0151" -o "E:\\WORKBUDDY\\清单.xlsx"
  python 文件夹提取工具.py -s "E:\\验货\\A0151" -o "D:\\导出" -n "我的清单"
  python 文件夹提取工具.py -s "E:\\验货\\A0151" -o "D:\\导出\\result.xlsx"
        """
    )

    parser.add_argument(
        "-s", "--source",
        type=str,
        help="要扫描的源目录路径（必填）"
    )

    parser.add_argument(
        "-o", "--output-dir",
        type=str,
        help="Excel输出目录路径（可选，默认为源目录同级）"
    )

    parser.add_argument(
        "-n", "--name",
        type=str,
        help="Excel文件名前缀（可选，不含.xlsx后缀）"
    )

    args = parser.parse_args()

    # 如果没有提供任何参数，进入交互模式
    if not args.source:
        print("=" * 60)
        print("  文件夹递归提取工具")
        print("=" * 60)
        print()
        print("使用说明：")
        print("  1. 输入要扫描的目录路径")
        print("  2. 输入Excel保存路径（默认保存到源目录同级）")
        print()

        # 交互式输入
        root_dir = input("📂 扫描目录: ").strip().strip('"').strip("'")
        if not root_dir:
            print("❌ 未输入目录路径")
            sys.exit(1)

        if not os.path.exists(root_dir):
            print(f"❌ 目录不存在: {root_dir}")
            sys.exit(1)

        if not os.path.isdir(root_dir):
            print(f"❌ 路径不是目录: {root_dir}")
            sys.exit(1)

        # 生成默认输出路径
        dir_name = os.path.basename(os.path.abspath(root_dir))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_save = os.path.join(os.path.dirname(root_dir),
                                    f"{dir_name}_文件夹清单_{timestamp}.xlsx")
        save_path = input(f"💾 Excel保存路径 [默认: {default_save}]: ").strip()
        if not save_path:
            save_path = default_save

    else:
        # 命令行模式
        root_dir = args.source
        output_dir = args.output_dir
        custom_name = args.name

        # 验证源目录
        if not os.path.exists(root_dir):
            print(f"❌ 目录不存在: {root_dir}")
            sys.exit(1)

        if not os.path.isdir(root_dir):
            print(f"❌ 路径不是目录: {root_dir}")
            sys.exit(1)

        # 生成输出路径
        if output_dir:
            # 指定了输出目录
            if custom_name:
                # 指定了自定义文件名
                if not custom_name.endswith('.xlsx'):
                    filename = custom_name + '.xlsx'
                else:
                    filename = custom_name
            else:
                # 默认文件名
                dir_name = os.path.basename(os.path.abspath(root_dir))
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"{dir_name}_文件夹清单_{timestamp}.xlsx"

            save_path = os.path.join(output_dir, filename)
            ensure_output_dir(save_path)

        elif custom_name:
            # 没指定输出目录，但指定了文件名 -> 保存到源目录同级
            dir_name = os.path.basename(os.path.abspath(root_dir))
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if custom_name.endswith('.xlsx'):
                filename = custom_name
            else:
                filename = f"{custom_name}_{timestamp}.xlsx"
            save_path = os.path.join(os.path.dirname(root_dir), filename)

        else:
            # 默认保存到源目录同级
            dir_name = os.path.basename(os.path.abspath(root_dir))
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            save_path = os.path.join(os.path.dirname(root_dir),
                                     f"{dir_name}_文件夹清单_{timestamp}.xlsx")

    # 确保扩展名正确
    if not save_path.lower().endswith('.xlsx'):
        save_path += '.xlsx'

    print(f"\n📂 扫描目录: {root_dir}")
    print(f"📁 输出文件: {save_path}")
    print()

    # 获取所有文件夹
    print("🔍 正在扫描...")
    all_dirs = get_all_folders(root_dir)

    if not all_dirs:
        print("⚠️  该目录下没有任何子文件夹")
        sys.exit(0)

    # 统计
    depth_stats = {}
    for d in all_dirs:
        depth_stats[d['depth']] = depth_stats.get(d['depth'], 0) + 1

    print(f"✅ 找到 {len(all_dirs)} 个文件夹")
    print()
    print("按层级分布:")
    for depth in sorted(depth_stats.keys()):
        print(f"  第 {depth} 层: {depth_stats[depth]} 个文件夹")
    print()

    # 生成Excel
    print("📊 正在生成Excel...")
    create_excel(all_dirs, save_path)

    print(f"\n✅ 完成! 共导出 {len(all_dirs)} 个文件夹")
    print(f"📁 文件位置: {save_path}\n")


if __name__ == "__main__":
    main()
