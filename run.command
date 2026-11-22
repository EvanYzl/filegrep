#!/bin/bash
# 双击启动 FileGrep 图形界面
cd "$(dirname "$0")" || exit 1

PY=".venv/bin/python"
if [ ! -x "$PY" ]; then
  echo "未找到虚拟环境，正在创建…"
  python3 -m venv --system-site-packages .venv || exit 1
  "$PY" -m pip install -q --upgrade pip
  "$PY" -m pip install -q PySide6-Essentials openpyxl python-docx python-pptx \
      pymupdf pdfplumber pypdf xlrd msoffcrypto-tool py7zr rarfile extract-msg \
      charset-normalizer markitdown olefile || exit 1
fi

exec "$PY" -m filegrep --gui "$@"
