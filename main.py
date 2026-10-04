# -*- coding: utf-8 -*-
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _missing_dep_error(exc: ImportError) -> None:
    """依赖缺失时给出可见提示（pythonw 下没有控制台，不提示就是静默退出）。"""
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "发票填表工具 - 缺少依赖库",
        f"缺少依赖：{exc.name or exc}\n\n"
        "解决办法（任选其一）：\n"
        "1. 双击本目录的 启动.bat（自动使用虚拟环境，推荐）\n"
        "2. 在本目录执行：\n"
        "   python -m pip install -r requirements.txt "
        "-i https://pypi.tuna.tsinghua.edu.cn/simple",
    )
    root.destroy()


try:
    from ui.app import main
except ImportError as e:
    _missing_dep_error(e)
    sys.exit(1)

if __name__ == "__main__":
    main()
