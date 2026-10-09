# -*- coding: utf-8 -*-
"""一键把本地交付版同步进开源仓库（release_开源发布/invoice-filler-1.0.0/）。

为什么需要它：
    本地版用「英文_中文备注」文件夹名（dist_安装包 / templates_模板配置 …），
    公开仓库保持纯英文（templates / tools / samples …）。两边**源码必须一致**，
    否则会各改各的、越漂越远。本脚本把"拷什么、改成什么名"这套映射固化成代码，
    以后发布只需跑这一条命令，不用手工复制，避免漏文件。

映射规则（只有这几处需要改名，其余目录根本不在仓库里）：
    templates_模板配置/  →  templates/
    tools_开发工具/      →  tools/
    示例模板/            →  samples/template/   （仓库侧目录）
    示例输出/            →  samples/output/     （仓库侧目录）

用法：
    python tools_开发工具/sync_to_repo.py            # 同步（会问你确认）
    python tools_开发工具/sync_to_repo.py --dry-run  # 只看要改什么，不落盘
    python tools_开发工具/sync_to_repo.py --yes      # 跳过确认（脚本化用）

同步完**不会自动提交**。要发布再自己 git add/commit/push。
"""
from __future__ import annotations

import argparse
import filecmp
import os
import shutil
import sys

# ---- 路径 ----
HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(HERE)                       # 项目根（本地交付版）
REPO_DIR = os.path.join(PROJECT_ROOT, "release_开源发布", "invoice-filler-1.0.0")

# ---- 重命名映射：本地相对路径 → 仓库相对路径 ----
# 目录级（整棵子树）
DIR_MAP = {
    "templates_模板配置": "templates",
    "tools_开发工具": "tools",
    "示例模板": os.path.join("samples", "template"),
    "示例输出": os.path.join("samples", "output"),
}

# 文件级：本地根目录下的文件 → 仓库相对路径（同名直接拷）
ROOT_FILES = [
    "main.py",
    "requirements.txt",
    "发票填表工具.spec",
    "启动.bat",
    "打包exe.bat",
    "README.md",
    ".gitignore",
]
# 注：README.md / .gitignore 由仓库自己维护，默认不覆盖（见 KEEP_REPO_FILES）

# 这些文件仓库自己有版本（README 的仓库版措辞不同），同步时默认**不动**
KEEP_REPO_FILES = {"README.md", ".gitignore"}

# 整个拷贝的源码目录
DIRS_TO_COPY = ["invoice_filler", "ui"]

# 模板目录只同步这一个：内置模板1 的种子。
# 其余 *.json 都是用户在本机导入的真实台账模板（source_file 指向他们的私人路径），
# 绝不能进公开仓库。
TEMPLATE_ALLOWLIST = {"template1.json"}

# 拷贝时忽略的杂物
IGNORE = shutil.ignore_patterns("__pycache__", "*.pyc", "*.pyo", ".DS_Store", "Thumbs.db")

# 禁止进入仓库的敏感目录/文件（真实客户数据）
FORBIDDEN = ["output_示例输出", "logs_运行日志", "build_构建缓存", "dist_安装包", ".venv"]


def _rel(p: str) -> str:
    return os.path.relpath(p, PROJECT_ROOT).replace("\\", "/")


def _same(a: str, b: str) -> bool:
    try:
        return filecmp.cmp(a, b, shallow=False)
    except OSError:
        return False


def collect_plan() -> list[tuple[str, str, str]]:
    """返回 [(动作, 源, 目标), ...]。动作 ∈ {copy, skip}。"""
    plan: list[tuple[str, str, str]] = []

    # 1) 源码包整体同步
    for d in DIRS_TO_COPY:
        src_dir = os.path.join(PROJECT_ROOT, d)
        if not os.path.isdir(src_dir):
            continue
        for root, _dirs, names in os.walk(src_dir):
            for n in names:
                if n.endswith((".pyc", ".pyo")) or "__pycache__" in root:
                    continue
                s = os.path.join(root, n)
                r = os.path.relpath(s, PROJECT_ROOT)
                plan.append(("copy", s, os.path.join(REPO_DIR, r)))

    # 2) 根目录文件
    for f in ROOT_FILES:
        s = os.path.join(PROJECT_ROOT, f)
        if not os.path.isfile(s):
            continue
        if f in KEEP_REPO_FILES:
            plan.append(("skip", s, os.path.join(REPO_DIR, f)))
        else:
            plan.append(("copy", s, os.path.join(REPO_DIR, f)))

    # 3) 改名目录
    for local_name, repo_rel in DIR_MAP.items():
        src_dir = os.path.join(PROJECT_ROOT, local_name)
        if not os.path.isdir(src_dir):
            continue
        allow = TEMPLATE_ALLOWLIST if local_name.startswith("templates") else None
        for root, _dirs, names in os.walk(src_dir):
            for n in names:
                if n.endswith((".pyc", ".pyo")) or "__pycache__" in root:
                    continue
                if allow is not None and n not in allow:
                    continue
                s = os.path.join(root, n)
                inner = os.path.relpath(s, src_dir)
                plan.append(("copy", s, os.path.join(REPO_DIR, repo_rel, inner)))

    return plan


def check_forbidden() -> list[str]:
    """确认敏感目录没被误纳入计划。"""
    bad = []
    repo_abs = os.path.abspath(REPO_DIR).lower()
    for f in FORBIDDEN:
        p = os.path.abspath(os.path.join(PROJECT_ROOT, f)).lower()
        if p.startswith(repo_abs) or repo_abs.startswith(p):
            bad.append(f)
    return bad


def sync(dry_run: bool = False) -> int:
    if not os.path.isdir(REPO_DIR):
        print(f"[!] 找不到仓库目录：{REPO_DIR}")
        return 1

    bad = check_forbidden()
    if bad:
        print(f"[!] 危险：以下目录与仓库目录重叠，已中止：{bad}")
        return 1

    plan = collect_plan()
    copied = skipped_same = kept = 0

    for action, src, dst in plan:
        if action == "skip":
            print(f"  [保留仓库版] {_rel(dst)}")
            kept += 1
            continue
        if os.path.exists(dst) and _same(src, dst):
            skipped_same += 1
            continue
        tag = "新增" if not os.path.exists(dst) else "更新"
        print(f"  [{tag}] {_rel(dst)}")
        if not dry_run:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
        copied += 1

    verb = "将同步" if dry_run else "已同步"
    print(f"\n{verb} {copied} 个文件（{skipped_same} 个内容相同跳过，{kept} 个保留仓库版）")
    if dry_run:
        print("（--dry-run：未写入任何文件）")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="把本地交付版同步进开源仓库")
    ap.add_argument("--dry-run", action="store_true", help="只显示要改什么，不落盘")
    ap.add_argument("--yes", action="store_true", help="跳过交互确认")
    args = ap.parse_args()

    print(f"本地源码：{PROJECT_ROOT}")
    print(f"开源仓库：{REPO_DIR}\n")

    if not args.dry_run and not args.yes:
        plan = collect_plan()
        changed = [p for p in plan if p[0] == "copy"
                   and not (os.path.exists(p[2]) and _same(p[1], p[2]))]
        print(f"即将同步 {len(changed)} 个文件到开源仓库（不会自动 git 提交）。")
        ans = input("继续？[y/N] ").strip().lower()
        if ans not in ("y", "yes"):
            print("已取消。")
            return 0

    return sync(dry_run=args.dry_run)


if __name__ == "__main__":
    sys.exit(main())
