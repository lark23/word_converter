### 1.工具描述

```markdown
# Word Converter 工具

**Word Converter** 是一个简单的命令行工具，旨在帮助用户将不同格式的文件（如 Word、PDF、PowerPoint、Excel）之间进行转换。该工具支持 Windows 系统，并依赖 `pywin32` 库来实现文档格式转换。

## 特性

- 支持将 Word 文件（.docx）转换为 PDF、Word、PPT、Excel 格式。
- 支持将 PDF 文件转换为 Word、PPT、Excel 格式。
- 支持将 PowerPoint 文件（.pptx）转换为 Word、PDF、Excel 格式。
- 支持将 Excel 文件（.xlsx）转换为 Word、PDF、PPT 格式。
- 使用命令行工具进行文件转换，简单易用。

## 安装

### 1. 克隆项目

首先，克隆该项目到本地：

```bash
git clone https://github.com/lark23/word_converter.git
cd word_converter
```

### 2. 安装依赖

确保你的 Python 环境已经安装，并且版本为 3.x。然后使用以下命令安装所有依赖：

```bash
pip install -r requirements.txt
```

### 3. 安装 `pywin32` 库

该工具依赖 `pywin32` 库来与 Microsoft Office 进行交互。你可以通过以下命令安装：

```bash
pip install pywin32
```

> 注意：`pywin32` 库仅支持 Windows 操作系统。

## 使用方法

### 1. 启动工具

在命令行中运行以下命令来启动工具并查看帮助信息：

```bash
python cli.py -h
```

### 2. 转换文件

启动工具后，您需要根据提示输入需要转换的文件路径，并选择目标文件格式。以下是命令行输入的一个示例：

```bash
==================================================
欢迎使用 Word 转换工具
这是一个简单的工具，用于将 Word 文件转换为 PDF、PPT、Excel 等格式。
==================================================
使用方法:
1. 请确保你已经有一个 Word 文件（.docx 格式）或其他支持的文件格式（如 PDF）。
2. 输入需要转换的文件路径，确保路径正确无误。
3. 输入目标输出文件的路径和格式。
4. 支持的输出格式有：pdf, word, ppt, excel。
5. 请确保输出路径有权限创建文件，避免写入错误。
==================================================
支持的转换格式：
1. Word 文件 (.docx) 可以转换为 PDF、Word、PPT、Excel。
2. PDF 文件可以转换为 Word、PPT、Excel。
3. PowerPoint 文件 (.pptx) 可以转换为 Word、PDF、Excel。
4. Excel 文件 (.xlsx) 可以转换为 Word、PDF、PPT。
==================================================
请根据提示输入需要转换的文件路径和目标格式。
请输入输入文件路径（支持 .docx, .pdf, .ppt, .xlsx 等格式）: D:\Documents\example.docx
请输入输出文件路径（支持 .pdf, .docx, .ppt, .xlsx 等格式）: D:\Documents\output.pdf
请选择目标格式（pdf、word、ppt、excel）: pdf

启动转换过程...
输入文件: D:\Documents\example.docx
输出文件: D:\Documents\output.pdf
目标格式: pdf

转换完成。
```

### 3. 参数说明

- `input_file`：输入文件的路径（支持格式：.docx、.pdf、.pptx、.xlsx 等）。
- `output_file`：输出文件的路径（支持格式：.pdf、.docx、.pptx、.xlsx 等）。
- `output_format`：目标文件格式，支持以下选项：
  - `pdf`：转换为 PDF 格式。
  - `word`：转换为 Word 格式（.docx）。
  - `ppt`：转换为 PowerPoint 格式（.pptx）。
  - `excel`：转换为 Excel 格式（.xlsx）。

## 支持的文件格式转换

| 输入文件格式       | 支持的输出格式               |
| ------------------ | ---------------------------- |
| Word (.docx)       | PDF, Word, PowerPoint, Excel |
| PDF                | Word, PowerPoint, Excel      |
| PowerPoint (.pptx) | Word, PDF, Excel             |
| Excel (.xlsx)      | Word, PDF, PowerPoint        |
