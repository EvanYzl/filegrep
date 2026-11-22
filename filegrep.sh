#!/bin/bash
# 命令行用法：./filegrep "正则" 目录 [更多参数]
cd "$(dirname "$0")" || exit 1
PY=".venv/bin/python"
[ -x "$PY" ] || PY="$(command -v python3)"
exec "$PY" -m filegrep "$@"
