# -*- coding: utf-8 -*-
"""路径与资源定位：兼容源码运行和 PyInstaller 打包后运行。"""
import os
import sys

# 内置模板配套的 Excel 源文件：释放到程序目录下，供「模板1」导出时沿用原版面。
# 模板 JSON 里记录的相对路径必须与它同名。
BUILTIN_TEMPLATE_WORKBOOK = "填表模板.xlsx"

# ---- 目录命名 ----
# 本地交付用「英文_中文备注」，一眼看清用途；公开仓库保持纯英文。
# 两种名字都兼容，所以同一份代码在两种目录结构下都能跑，同步到仓库时不必改代码。
TEMPLATE_DIR_NAMES = ("templates_模板配置", "templates")
LOGS_DIR_NAMES = ("logs_运行日志", "logs")


def app_base_dir() -> str:
    """应用程序根目录：打包后为 exe 所在目录，源码运行为项目根目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _package_dir() -> str:
    """invoice_filler 包所在目录（内置资源的存放位置）。"""
    return os.path.dirname(os.path.abspath(__file__))


def _is_writable_dir(d: str) -> bool:
    """目录可写？（不存在则尝试创建）"""
    try:
        os.makedirs(d, exist_ok=True)
        probe = os.path.join(d, ".write_test")
        with open(probe, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(probe)
        return True
    except OSError:
        return False


def _prefers_english_names() -> bool:
    """判断当前是"公开仓库结构"还是"本地交付结构"。

    依据：程序目录里若已存在英文名目录（templates/ 或 tools/），
    说明这是一份仓库代码；本地交付版只会出现中文备注名。
    → 仓库版下新建目录也用英文名，免得仓库里冒出一个中文目录。
    """
    base = app_base_dir()
    return any(os.path.isdir(os.path.join(base, n)) for n in ("templates", "tools"))


def _pick_dir(names) -> str:
    """在程序目录下挑一个可写目录。

    优先级：
      1. 已存在的可写目录按 names 顺序（本地中文名在前，命中即用）；
      2. 若都不存在，按当前结构决定新建哪个名字——
         仓库结构（已有 templates/ 或 tools/）建英文名，本地交付版建中文名。
    """
    base = app_base_dir()
    candidates = [os.path.join(base, n) for n in names]
    for c in candidates:              # 1. 已存在的目录优先沿用
        if os.path.isdir(c) and _is_writable_dir(c):
            return c
    ordered = list(candidates)        # 2. 都不存在 → 按结构选命名
    if _prefers_english_names():
        ordered = ordered[::-1]       # names 里英文名排在后，倒过来即优先英文
    for c in ordered:
        if _is_writable_dir(c):
            return c
    return ""


def _user_dir_fallback(sub: str) -> str:
    d = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "发票填表工具", sub)
    os.makedirs(d, exist_ok=True)
    return d


def templates_dir() -> str:
    """模板 JSON 存放目录。程序目录下的 templates_模板配置/（公开仓库版为 templates/），
    若都不可写则回退到用户目录，保证新导入的模板仍能保存。"""
    return _pick_dir(TEMPLATE_DIR_NAMES) or _user_dir_fallback("templates")


def logs_dir() -> str:
    """运行日志目录（拖放链路日志）。"""
    return _pick_dir(LOGS_DIR_NAMES) or _user_dir_fallback("logs")


def resolve_template_source(source_file: str) -> str:
    """把模板记录的来源解析成可用的绝对路径；找不到时返回空串。

    模板 JSON 里的 source_file 有两种合法写法：
    - 绝对路径：用户自己「从 Excel 导入」的模板，指向他们自己的台账文件；
    - 相对路径：随程序分发的内置模板（如「填表模板.xlsx」），相对**程序目录**解析——
      源码运行时即项目根目录，exe 运行时即 exe 所在目录。

    这样同一份模板 JSON 在本机开发和别人电脑上都能定位到模板原件，
    不会因为写死了某一台机器的绝对路径而静默退化成"重建的素表"。
    """
    if not source_file:
        return ""
    if os.path.isabs(source_file):
        return source_file if os.path.exists(source_file) else ""
    cand = os.path.join(app_base_dir(), source_file)
    return cand if os.path.exists(cand) else ""


def _copy_if_missing(src: str, dst: str) -> bool:
    """仅在目标不存在时复制；失败（如只读目录）静默跳过，不阻断启动。"""
    import shutil

    try:
        if os.path.exists(dst) or not os.path.exists(src):
            return False
        parent = os.path.dirname(dst)
        if parent:
            os.makedirs(parent, exist_ok=True)
        shutil.copyfile(src, dst)
        return True
    except OSError:
        return False


def _migrate_stale_template_sources() -> None:
    """老版本的 template1.json 里记的是开发机绝对路径（如 E:\\自制程序\\...），
    换到别人电脑必然定位不到。若同名相对路径可用，就地改写成相对路径。"""
    import json

    d = templates_dir()
    try:
        names = [n for n in os.listdir(d) if n.endswith(".json")]
    except OSError:
        return
    for fn in names:
        p = os.path.join(d, fn)
        try:
            with open(p, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except (OSError, ValueError):
            continue
        if not isinstance(cfg, dict):
            continue
        src = cfg.get("source_file") or ""
        if not src or resolve_template_source(src):
            continue
        rel = os.path.basename(src)
        # 只迁移"指向内置模板原件"的那一条，避免误改用户自己的模板路径
        if rel != BUILTIN_TEMPLATE_WORKBOOK or not resolve_template_source(rel):
            continue
        cfg["source_file"] = rel
        try:
            with open(p, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except OSError:
            pass


def ensure_builtin_template() -> None:
    """首次运行时释放内置模板：配置 template1.json + 配套的 填表模板.xlsx。

    xlsx 刻意释放到**程序目录**（与模板 JSON 里的相对路径一致）：
    源码运行时是项目根目录，exe 运行时是 exe 旁边——
    两种情况下导出都能沿用原模板的表头样式与工作表结构。
    """
    _copy_if_missing(os.path.join(_package_dir(), "builtin_template1.json"),
                     os.path.join(templates_dir(), "template1.json"))
    _copy_if_missing(os.path.join(_package_dir(), "builtin_template.xlsx"),
                     os.path.join(app_base_dir(), BUILTIN_TEMPLATE_WORKBOOK))
    _migrate_stale_template_sources()
