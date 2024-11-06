# -- coding: utf-8 --
from word_to_pdf import convert_word_to_pdf
from pdf_to_other import pdf_to_word, pdf_to_ppt, pdf_to_excel

def convert(input_file, output_file, output_format):
    """
    统一入口：根据文件格式进行转换。
    :param input_file: 输入文件路径
    :param output_file: 输出文件路径
    :param output_format: 输出格式，'pdf', 'word', 'ppt', 'excel'
    """
    if output_format == "pdf":
        if input_file.lower().endswith(".docx"):
            convert_word_to_pdf(input_file, output_file)
            print(f"Word 转 PDF 成功，保存为 {output_file}")
        else:
            print("只支持 .docx 文件转换为 PDF。")
    elif output_format == "word":
        if input_file.lower().endswith(".pdf"):
            pdf_to_word(input_file, output_file)
            print(f"PDF 转 Word 成功，保存为 {output_file}")
        else:
            print("只支持 PDF 文件转换为 Word。")
    elif output_format == "ppt":
        if input_file.lower().endswith(".pdf"):
            pdf_to_ppt(input_file, output_file)
            print(f"PDF 转 PPT 成功，保存为 {output_file}")
        else:
            print("只支持 PDF 文件转换为 PPT。")
    elif output_format == "excel":
        if input_file.lower().endswith(".pdf"):
            pdf_to_excel(input_file, output_file)
            print(f"PDF 转 Excel 成功，保存为 {output_file}")
        else:
            print("只支持 PDF 文件转换为 Excel。")
    else:
        print("不支持的输出格式。")
