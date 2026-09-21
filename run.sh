#!/usr/bin/env bash
# macOS / Linux 启动脚本
set -e
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$DIR"

VENV="$DIR/venv"
if [ ! -d "$VENV" ]; then
  echo "🔧 首次运行，创建虚拟环境..."
  python3 -m venv "$VENV"
  "$VENV/bin/pip" install --upgrade pip
  "$VENV/bin/pip" install -r requirements.txt
fi

echo "🚀 启动商务英语毕业论文助手..."
"$VENV/bin/python3" "$DIR/app.py"
