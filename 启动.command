#!/bin/bash
# 双击这个文件即可启动毕业论文助手
cd "$(dirname "$0")"

VENV_PY="./venv/bin/python"
SYS_HB_PY="/opt/homebrew/bin/python3"
SYS_PY="/usr/bin/python3"

# 优先用项目自带 venv（依赖齐全）
if [ -x "$VENV_PY" ] && $VENV_PY -c "import _tkinter, docx, openpyxl" 2>/dev/null; then
    exec $VENV_PY app.py
fi

# 次选 Homebrew Python + 自动装依赖
if [ -x "$SYS_HB_PY" ] && $SYS_HB_PY -c "import _tkinter" 2>/dev/null; then
    $SYS_HB_PY -c "import docx, openpyxl" 2>/dev/null || {
        osascript -e 'display alert "缺少依赖" message "正在为你安装 python-docx 和 openpyxl，请稍候..."'
        $SYS_HB_PY -m pip install --break-system-packages --user python-docx openpyxl 2>&1 | tail -5
    }
    exec $SYS_HB_PY app.py
fi

# 兜底：系统 Python
if [ -x "$SYS_PY" ] && $SYS_PY -c "import _tkinter" 2>/dev/null; then
    osascript -e 'display alert "提示" message "你的 macOS 系统 Python 的 tk 渲染有问题。建议运行：brew install python-tk@3.14"'
    $SYS_PY -m pip install --user python-docx openpyxl 2>/dev/null
    exec $SYS_PY app.py
fi

osascript -e 'display alert "无法启动" message "找不到可用的 Python。请在终端运行：brew install python-tk@3.14"'
exit 1
