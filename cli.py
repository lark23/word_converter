# -- coding: utf-8 --
# cli.py
import argparse
from word_converter.converter import convert


def main():
    # 设置欢迎信息和描述
    print("=" * 50)
    print("欢迎使用 Word 转换工具")
    print("这是一个简单的工具，用于将 Word 文件转换为 PDF、PPT、Excel 等格式。")
    print("=" * 50)

    # 显示如何使用该工具
    print("使用方法:")
    print("1. 请确保你已经有一个 Word 文件（.docx 格式）或其他支持的文件格式（如 PDF）。")
    print("2. 输入需要转换的文件路径，确保路径正确无误。")
    print("3. 输入目标输出文件的路径和格式。")
    print("4. 支持的输出格式有：pdf, word, ppt, excel。")
    print("5. 请确保输出路径有权限创建文件，避免写入错误。")
    print("=" * 50)

    # 添加可转换格式的提示
    print("支持的转换格式：")
    print("1. Word 文件 (.docx) 可以转换为 PDF、Word、PPT、Excel。")
    print("2. PDF 文件可以转换为 Word、PPT、Excel。")
    print("3. PowerPoint 文件 (.pptx) 可以转换为 Word、PDF、Excel。")
    print("4. Excel 文件 (.xlsx) 可以转换为 Word、PDF、PPT。")
    print("=" * 50)

    # 提示用户接下来需要输入的信息
    print("请根据提示输入需要转换的文件路径和目标格式。")

    # 输入文件路径和目标格式
    input_file = input("请输入输入文件路径（支持 .docx, .pdf, .ppt, .xlsx 等格式）: ")
    output_file = input("请输入输出文件路径（支持 .pdf, .docx, .ppt, .xlsx 等格式）: ")

    # 提示选择目标格式并接收用户输入
    output_format = input("请选择目标格式（pdf、word、ppt、excel）: ").lower()

    # 检查目标格式是否合法
    if output_format not in ['pdf', 'word', 'ppt', 'excel']:
        print("无效的目标格式，请选择 'pdf'、'word'、'ppt' 或 'excel'。")
        return

    # 显示用户输入的信息
    print(f"\n启动转换过程...")
    print(f"输入文件: {input_file}")
    print(f"输出文件: {output_file}")
    print(f"目标格式: {output_format}")

    # 调用转换函数
    try:
        convert(input_file, output_file, output_format)
        print("\n转换完成。")
    except Exception as e:
        print(f"转换过程中出现错误: {e}")


if __name__ == "__main__":
    main()


