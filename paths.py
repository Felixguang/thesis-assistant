"""
统一路径处理：支持开发模式 + PyInstaller 单文件打包模式。

- resource_path(): 读资源（开发模式用 __file__；打包后用 sys._MEIPASS）
- user_data_path(): 写用户数据（永远写到 ~/Documents/毕业论文助手/）
- atomic_write(): 原子写入（写临时文件后 os.replace）
"""
import os
import sys
import tempfile
from pathlib import Path


def resource_path(rel_path: str) -> Path:
    """获取资源绝对路径。开发模式用 __file__ 目录；PyInstaller 单文件用 sys._MEIPASS。

    示例:
        resource_path("data/topics.json")
        # 开发模式: /path/to/_prototype/data/topics.json
        # 打包后:   /var/folders/xxx/_MEIXXXXX/data/topics.json
    """
    base = getattr(sys, "_MEIPASS", Path(__file__).parent)
    return Path(base) / rel_path


def user_data_dir() -> Path:
    """用户数据目录（可写）。macOS/Linux/Windows 一致写到 ~/Documents/毕业论文助手/。"""
    data_dir = Path.home() / "Documents" / "毕业论文助手"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir


def user_data_path(filename: str) -> Path:
    """用户数据文件路径。首次访问时自动创建目录。"""
    return user_data_dir() / filename


def atomic_write(target: Path, content: str, encoding: str = "utf-8") -> None:
    """原子写入：先写临时文件，成功后 os.replace 覆盖。

    避免写入过程中崩溃导致半截文件污染数据。
    """
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(
        dir=str(target.parent),
        prefix=f".{target.name}.",
        suffix=".tmp",
    )
    try:
        with os.fdopen(fd, "w", encoding=encoding) as f:
            f.write(content)
        os.replace(tmp_path, target)
    except Exception:
        # 清理临时文件
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise