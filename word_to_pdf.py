# -- coding: utf-8 --
import os
from win32com import client

def convert_word_to_pdf(word_file, pdf_file):
    # 创建 Word 应用实例
    word = client.Dispatch("Word.Application")
    word.Visible = False
    # 打开 Word 文件
    doc = word.Documents.Open(word_file)
    # 保存为 PDF
    doc.SaveAs(pdf_file, FileFormat=17)  # 17 表示 PDF 格式
    # 关闭文档
    doc.Close()
    # 退出 Word 应用
    word.Quit()
