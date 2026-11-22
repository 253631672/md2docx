#!/bin/bash
# 一条命令跑通：把 demo.md 转成 docx，再转一份横向版
set -e
cd "$(dirname "$0")"

python3 ../md2docx.py demo.md --title "门店周报" --author "Marco" -o demo.docx
python3 ../md2docx.py demo.md --landscape --size 9 -o demo_横向.docx

ls -lh demo.docx demo_横向.docx
echo "用 Word/WPS 打开 demo.docx 检查效果"
