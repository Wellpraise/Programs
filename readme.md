# 文件分类识别程序 - Windows 环境部署记录

## 1. 程序功能
该程序用于自动扫描指定目录下的 PDF、图片及压缩包，通过 OCR 技术识别其中的“提单号”或“海关单号”，并根据设定的规则自动创建文件夹进行分类存放，最后生成一份 Excel 汇总清单。

## 2. 运行环境依赖 (必须安装)

为了让程序正常工作，电脑必须安装以下三个系统级工具，并配置环境变量 `Path`。

### (1) Poppler (PDF 转图片工具)
- **作用**: 将 PDF 页面转换为图像，以便 OCR 识别。
- **安装路径**: `E:\Coprograms\Library\bin` (以实际解压路径为准)
- **配置**: 将上述 `bin` 文件夹路径添加到系统环境变量 `Path` 中。
- **验证命令**: `pdftoppm -h`

### (2) Tesseract OCR (文字识别引擎)
- **作用**: 识别图片中的中文和英文文字。
- **安装要点**: 安装时必须在 `Additional language data` 中勾选 **Han Simplified (简体中文)** 和 **English**。
- **安装路径**: `C:\Program Files\Tesseract-OCR`
- **配置**: 将该安装路径添加到系统环境变量 `Path` 中。
- **验证命令**: `tesseract --version`

### (3) 7-Zip (全能解压工具)
- **作用**: 自动解压 `.zip`, `.rar`, `.7z` 等压缩包。
- **安装路径**: `D:\7z\7z.exe`
- **配置**: 程序代码中已硬编码该路径，无需添加环境变量。

---

## 3. Python 依赖库安装
运行以下命令安装必要的 Python 插件：
```bash
pip install pandas openpyxl pdfplumber pytesseract pillow pdf2image rarfile
4. 目录结构配置
程序设定在 E:\Coprograms 目录下运行，结构如下：
E:\Coprograms\fileclassifier.py : 程序主文件
E:\Coprograms\input\ : 【输入区】 放入待处理的 PDF/图片/压缩包
E:\Coprograms\output\ : 【输出区】 存放分类后的文件夹和 Excel 清单
E:\Coprograms\temp_unzip\ : 【临时区】 自动解压缓存，程序运行完会自动清理

5. 运行方式
打开终端（PowerShell 或 CMD），执行：

python E:\Coprograms\fileclassifier.py


6.常见问题记录 (Troubleshooting)
报错 D:\test 不存在: 原因是代码中的路径未修改。需确保 ROOT_DIR 等路径指向 E 盘。
报错 Cannot find working tool: 之前尝试使用 rarfile 依赖 WinRAR，后升级为直接调用 7z.exe 解决。
文件名 .yp vs .py: 请确保文件名后缀为 .py，否则 Python 解释器可能无法正确识别。