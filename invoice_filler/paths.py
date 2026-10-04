# -*- coding: utf-8 -*-
"""路径与资源定位：兼容源码运行和 PyInstaller 打包后运行。"""
import os
import sys


def app_base_dir() -> str:
    """应用程序根目录：打包后为 exe 所在目录，源码运行为项目根目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def templates_dir() -> str:
    """模板 JSON 存放目录。优先使用程序目录下的 templates/（便携式），
    若不可写则回退到用户目录，保证新导入的模板可以保存。"""
    local = os.path.join(app_base_dir(), "templates")
    try:
        os.makedirs(local, exist_ok=True)
        probe = os.path.join(local, ".write_test")
        with open(probe, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(probe)
        return local
    except OSError:
        fallback = os.path.join(
            os.environ.get("APPDATA", os.path.expanduser("~")), "发票填表工具", "templates"
        )
        os.makedirs(fallback, exist_ok=True)
        return fallback


def ensure_builtin_template() -> None:
    """首次运行时把内置的模板1释放到模板目录。"""
    import shutil

    dst = os.path.join(templates_dir(), "template1.json")
    if os.path.exists(dst):
        return
    src = os.path.join(os.path.dirname(os.path.abspath(__file__)), "builtin_template1.json")
    if os.path.exists(src):
        shutil.copyfile(src, dst)
