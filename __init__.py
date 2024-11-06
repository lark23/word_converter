# -- coding: utf-8 --
# __init__.py

# 设定包的版本
__version__ = "0.1.0"

# 从模块导入功能
from .word_to_pdf import convert_word_to_pdf
from .pdf_to_other import pdf_to_word, pdf_to_ppt, pdf_to_excel
from .converter import convert
