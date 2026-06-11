#!/bin/bash
# macOS 双击运行：提取本文件夹内所有 PDF 发票，生成汇总 Excel
# 使用前提：已安装 Python 3 和依赖（pip3 install pdfplumber openpyxl），
# 且本文件与 extract_invoices.py、发票 PDF 在同一个文件夹
cd "$(dirname "$0")"
python3 extract_invoices.py "$PWD"
echo ""
read -p "按回车键关闭窗口..."
