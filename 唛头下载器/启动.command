#!/bin/bash
cd "$(dirname "$0")"
if [ ! -f .venv/bin/python ]; then
  echo "首次运行，正在创建虚拟环境并安装依赖…"
  python3 -m venv .venv
  .venv/bin/pip install --quiet flask requests openpyxl
fi
(sleep 1.5 && open http://127.0.0.1:8756) &
.venv/bin/python app.py
