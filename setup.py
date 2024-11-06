# -- coding: utf-8 --
from setuptools import setup, find_packages

# 读取 README.md 内容，作为 long_description
with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="word_converter",  # 包名称
    version="0.1.0",  # 包版本
    author="您的名字",  # 作者
    author_email="您的邮箱",  # 作者邮箱
    description="一个简单的 Python 包，支持将 Word 文件转换为 PDF，并支持 PDF 转换为 Word、PPT、Excel 等格式",  # 简短描述
    long_description=long_description,  # 从 README.md 获取详细描述
    long_description_content_type="text/markdown",  # 说明 README.md 使用 Markdown 格式
    url="https://github.com/yourusername/word_converter",  # 包的主页
    packages=find_packages(),  # 自动发现包中的所有模块
    classifiers=[
        "Programming Language :: Python :: 3",  # Python 3.x 支持
        "License :: OSI Approved :: MIT License",  # 许可协议
        "Operating System :: Microsoft :: Windows",  # 仅支持 Windows 系统
        "Programming Language :: Python :: 3.8",  # Python 版本
        "Programming Language :: Python :: 3.9",
    ],
    python_requires=">=3.6",  # 要求 Python 版本至少为 3.6
    install_requires=[  # 依赖库
        "pywin32==303",
        "pdf2docx==0.5.5",
        "openpyxl==3.0.10",
        "python-pptx==0.6.21",
        "argparse==1.4.0"
    ],
    entry_points={  # 配置命令行工具
        'console_scripts': [
            'word-converter=cli:main',  # 指定命令行入口点
        ],
    },
)
