# -*- coding: utf-8 -*-
"""发票填表工具 - 桌面界面（tkinter）。

功能：批量导入发票文件（PDF/图片/扫描件）→ 后台识别（进度不卡界面）→
预览结果 → 双击人工修正 → 按模板导出 Excel；模板管理与字段映射说明。
"""
from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from invoice_filler import __version__, paths
from invoice_filler.exporter import export
from invoice_filler.fields import ALL_FIELDS
from invoice_filler.models import STATUS_TEXT, InvoiceData
from invoice_filler.parser import parse_invoice
from invoice_filler.report import generate_mapping_doc
from invoice_filler.template import (TemplateConfig, build_template_from_workbook,
                                     load_templates, match_quality)

SUPPORTED_EXT = (".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")

# ---- 可选的拖拽支持（tkinterdnd2，未安装时程序仍可用，只是不能拖） ----
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _probe = TkinterDnD.Tk()
    _probe.withdraw()
    _probe.destroy()
    _DND = True
    BaseTk = TkinterDnD.Tk
except Exception:
    DND_FILES = None
    _DND = False
    BaseTk = tk.Tk


def scan_folder(folder: str) -> list[str]:
    """递归收集文件夹内支持的发票文件。"""
    found = []
    for root, _dirs, names in os.walk(folder):
        for n in sorted(names):
            if n.lower().endswith(SUPPORTED_EXT):
                found.append(os.path.join(root, n))
    return found


class EditDialog(tk.Toplevel):
    """单元格人工修正小窗。"""

    def __init__(self, app: "InvoiceFillerApp", row_iid: str, col_key: str,
                 col_header: str, value: str):
        super().__init__(app)
        self.app = app
        self.row_iid = row_iid
        self.col_key = col_key
        self.title(f"修正「{col_header}」")
        self.resizable(False, False)
        self.transient(app)
        self.grab_set()
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=f"{col_header}（留空表示清除修正，恢复识别值/留白）").pack(anchor="w")
        self.var = tk.StringVar(value=value)
        entry = ttk.Entry(frm, textvariable=self.var, width=48)
        entry.pack(fill="x", pady=(4, 8))
        entry.select_range(0, "end")
        btns = ttk.Frame(frm)
        btns.pack(fill="x")
        ttk.Button(btns, text="确定", command=self._ok).pack(side="right", padx=2)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="right", padx=2)
        entry.bind("<Return>", lambda e: self._ok())
        entry.bind("<Escape>", lambda e: self.destroy())
        entry.focus_set()
        x = app.winfo_pointerx() - 120
        y = app.winfo_pointery() - 20
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _ok(self):
        self.app.commit_edit(self.row_iid, self.col_key, self.var.get())
        self.destroy()


class TemplateManager(tk.Toplevel):
    """模板管理：切换 / 从 Excel 导入新模板 / 删除。"""

    def __init__(self, app: "InvoiceFillerApp"):
        super().__init__(app)
        self.app = app
        self.title("模板管理")
        self.geometry("520x360")
        self.transient(app)
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="已有模板（每个模板是一份 *.json，保存在下面这个目录里，可随程序一起复制分享）：").pack(anchor="w")
        loc = ttk.Frame(frm)
        loc.pack(fill="x", pady=(2, 6))
        ttk.Label(loc, text=paths.templates_dir(), foreground="#555555").pack(side="left")
        ttk.Button(loc, text="打开该文件夹", command=self.open_templates_dir).pack(side="right")
        self.listbox = tk.Listbox(frm, height=10)
        self.listbox.pack(fill="both", expand=True, pady=(0, 6))
        self.detail_var = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self.detail_var, foreground="#555555",
                  wraplength=480, justify="left").pack(anchor="w", pady=(0, 4))
        self.reload_list()
        btns = ttk.Frame(frm)
        btns.pack(fill="x")
        ttk.Button(btns, text="设为当前模板", command=self.use_selected).pack(side="left", padx=2)
        ttk.Button(btns, text="从 Excel 导入新模板…", command=self.import_template).pack(side="left", padx=2)
        ttk.Button(btns, text="删除", command=self.delete_selected).pack(side="left", padx=2)
        ttk.Button(btns, text="关闭", command=self.destroy).pack(side="right", padx=2)
        self.listbox.bind("<Double-Button-1>", lambda e: self.use_selected())
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

    def open_templates_dir(self):
        d = paths.templates_dir()
        try:
            os.makedirs(d, exist_ok=True)
            os.startfile(d)  # noqa: S606
        except OSError as e:
            messagebox.showerror("无法打开文件夹", f"{d}\n\n{e}", parent=self)

    def _on_select(self, _e=None):
        sel = self.listbox.curselection()
        if not sel or sel[0] >= len(self.app.templates):
            self.detail_var.set("")
            return
        t = self.app.templates[sel[0]]
        src = t.source_file or "（无）"
        if src and src != "（无）":
            resolved = paths.resolve_template_source(src)
            src = resolved or f"{src}  ← 已失效，导出将按配置重建表格"
        self.detail_var.set(f"模板来源文件：{src}")

    def reload_list(self):
        self.app.reload_templates()
        self.listbox.delete(0, "end")
        for t in self.app.templates:
            cur = self.app.current_template.id if self.app.current_template else None
            mark = "（当前）" if t.id == cur else ""
            self.listbox.insert("end", f"{t.name}  [{t.id}]  {len(t.columns)}列 {mark}")
        if self.app.templates:
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set(0)
            self._on_select()

    def _selected(self) -> TemplateConfig | None:
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showinfo("提示", "请先选择一个模板", parent=self)
            return None
        return self.app.templates[sel[0]]

    def use_selected(self):
        t = self._selected()
        if t:
            self.app.set_current_template(t)
            self.reload_list()

    def delete_selected(self):
        t = self._selected()
        if not t:
            return
        if t.id == self.app.current_template.id:
            messagebox.showwarning("提示", "不能删除当前使用的模板", parent=self)
            return
        if not messagebox.askyesno("确认", f"删除模板「{t.name}」？", parent=self):
            return
        p = os.path.join(paths.templates_dir(), f"{t.id}.json")
        if os.path.exists(p):
            os.remove(p)
        self.reload_list()

    def import_template(self):
        path = filedialog.askopenfilename(
            parent=self, title="选择含表头的 Excel 模板",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if not path:
            return
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
        try:
            sheets = wb.sheetnames
        finally:
            wb.close()
        sheet = sheets[0]
        if len(sheets) > 1:
            sheet = SheetPicker(self, sheets).choice
            if not sheet:
                return
        try:
            tpl = build_template_from_workbook(path, sheet=sheet)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("导入失败", f"读取模板失败：{e}", parent=self)
            return
        try:
            editor = MappingEditor(self, tpl)
        except Exception as e:  # noqa: BLE001 - 构建失败不能留下半成品窗口
            messagebox.showerror("导入失败", f"无法打开映射确认窗口：\n{type(e).__name__}: {e}",
                                 parent=self)
            return
        self.wait_window(editor)
        if editor.saved:
            self.reload_list()


class SheetPicker(tk.Toplevel):
    """选择工作表。"""

    def __init__(self, parent, sheets):
        super().__init__(parent)
        self.title("选择工作表")
        self.choice: str | None = None
        self.transient(parent)
        self.grab_set()
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="该 Excel 含多个工作表，请选择表头所在的工作表：").pack(anchor="w")
        lb = tk.Listbox(frm, height=6)
        lb.pack(fill="both", expand=True, pady=6)
        for s in sheets:
            lb.insert("end", s)
        lb.selection_set(0)

        def ok():
            sel = lb.curselection()
            self.choice = sheets[sel[0]] if sel else None
            self.destroy()

        lb.bind("<Double-Button-1>", lambda e: ok())
        ttk.Button(frm, text="确定", command=ok).pack(side="right")
        x, y = parent.winfo_rootx() + 80, parent.winfo_rooty() + 120
        self.geometry(f"+{x}+{y}")
        self.wait_window()


class MappingEditor(tk.Toplevel):
    """导入新模板后确认/调整表头映射。"""

    def __init__(self, parent, template: TemplateConfig):
        super().__init__(parent)
        self.template = template
        self.saved = False
        self.title(f"确认映射 - {template.name}")
        self.geometry("760x520")
        self.minsize(640, 420)
        self.transient(parent)
        self.grab_set()

        q = match_quality(template)
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill="x", side="top")
        ttk.Label(top, text=(f"表头共 {q['total']} 列：自动匹配 {q['mapped']} 列，"
                             f"{q['unmapped']} 列无匹配（默认留白）。"
                             "可在下方下拉框人工调整，调整后请保存。")).pack(anchor="w")
        ttk.Label(top, text=f"保存位置：{paths.templates_dir()}",
                  foreground="#555555").pack(anchor="w", pady=(2, 0))

        # 按钮栏必须"先"占位（side="bottom"），否则会被下面撑满的 canvas 挤成一条缝
        btns = ttk.Frame(self, padding=(10, 8))
        btns.pack(side="bottom", fill="x")
        ttk.Button(btns, text="保存为新模板", command=self._save).pack(side="right", padx=4)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="right")

        mid = ttk.Frame(self)
        mid.pack(fill="both", expand=True, side="top")
        canvas = tk.Canvas(mid, highlightthickness=0)
        sb = ttk.Scrollbar(mid, orient="vertical", command=canvas.yview)
        body = ttk.Frame(canvas, padding=(10, 4))
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=body, anchor="nw", tags="body")
        # 内层 frame 宽度跟随画布。注意：Canvas.tag_bind 只接受键鼠类事件，
        # 绑 <Configure> 会抛 TclError —— 旧版正崩在这里，导致窗口只建了一半，
        # 下拉映射框和「保存为新模板」按钮压根没被创建出来。必须绑在画布控件本身。
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure("body", width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.choices: list[tuple[object, ttk.Combobox]] = []
        options = ["（留白）"] + [f["label"] for f in ALL_FIELDS.values()]
        label_to_key = {"（留白）": None}
        key_to_label = {}
        for f in ALL_FIELDS.values():
            label_to_key[f["label"]] = f["key"]
            key_to_label[f["key"]] = f["label"]

        for i, col in enumerate(template.columns):
            row = ttk.Frame(body)
            row.grid(row=i, column=0, columnspan=3, sticky="we", padx=2, pady=2)
            span = f"{col.col}" if col.colspan <= 1 else f"{col.col}起{col.colspan}列"
            ttk.Label(row, text=f"{span}", width=9).pack(side="left")
            ttk.Label(row, text=col.header, width=22).pack(side="left")
            cb = ttk.Combobox(row, values=options, state="readonly", width=24)
            cb.set(key_to_label.get(col.field, "（留白）"))
            cb.pack(side="left", padx=6)
            self.choices.append((col, cb))
        body.columnconfigure(0, weight=1)

    def _save(self):
        name = simpledialog.askstring("模板名称", "给新模板起个名字：",
                                      initialvalue=self.template.name, parent=self)
        if not name:
            return
        self.template.name = name
        label_to_key = {"（留白）": None}
        for k, f in ALL_FIELDS.items():
            label_to_key[f["label"]] = k
        for col, cb in self.choices:
            col.field = label_to_key.get(cb.get())
            f = ALL_FIELDS.get(col.field) if col.field else None
            col.note = f["source"] if f else "发票中无对应字段，留白"
        try:
            path = self.template.save()
        except OSError as e:
            messagebox.showerror(
                "保存失败",
                f"无法写入模板目录：\n{paths.templates_dir()}\n\n{e}\n\n"
                "请确认该目录可写，或换到有写权限的位置运行程序。", parent=self)
            return
        # 立刻回读一次，确保"保存了"不是错觉——下次启动真的能加载出来
        try:
            ok = TemplateConfig.load(path).id == self.template.id
        except Exception:  # noqa: BLE001
            ok = False
        if not ok:
            messagebox.showerror(
                "保存失败", f"模板已写入但无法回读，请检查目录权限：\n{path}", parent=self)
            return
        self.saved = True
        q = match_quality(self.template)
        messagebox.showinfo(
            "模板已保存",
            f"模板「{name}」已保存，关闭程序后依然保留。\n\n"
            f"存储位置：\n{path}\n\n"
            f"它已出现在「模板」下拉框和模板管理列表里，可直接切换使用。\n"
            f"已映射 {q['mapped']} 列，留白 {q['unmapped']} 列。",
            parent=self)
        self.destroy()


class InvoiceFillerApp(BaseTk):
    def __init__(self):
        super().__init__()
        self.title(f"发票填表工具 v{__version__}")
        self.geometry("1180x720")
        self.minsize(900, 560)
        self.option_add("*Font", ("Microsoft YaHei UI", 10))

        self.templates: list[TemplateConfig] = []
        self.current_template: TemplateConfig | None = None
        self.files: list[str] = []
        self.records: dict[int, InvoiceData] = {}
        self.queue: queue.Queue = queue.Queue()
        self.running = False

        self._build_ui()
        if _DND:
            self._enable_drag_drop()
            if self._is_elevated():
                self.status_var.set("警告：程序以管理员身份运行时 Windows 会禁止从资源管理器拖放，请用普通方式启动")
        else:
            self.status_var.set("未安装 tkinterdnd2，拖拽不可用（pip install tkinterdnd2）；请用按钮导入文件")
        self.reload_templates()
        self.after(120, self._poll_queue)

    @staticmethod
    def _is_elevated() -> bool:
        """Windows 下检测当前进程是否以管理员权限运行（此时资源管理器拖放被系统禁止）。"""
        if sys.platform != "win32":
            return False
        try:
            import ctypes
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return False

    # ---------------- UI 构建 ----------------
    def _build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:
            pass

        bar = ttk.Frame(self, padding=(8, 6))
        bar.pack(fill="x")
        ttk.Button(bar, text="导入发票文件…", command=self.add_files_dialog).pack(side="left", padx=2)
        ttk.Button(bar, text="导入整个文件夹", command=self.add_folder_dialog).pack(side="left", padx=2)
        ttk.Button(bar, text="移除选中", command=self.remove_selected_files).pack(side="left", padx=2)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Label(bar, text="模板：").pack(side="left")
        self.tpl_var = tk.StringVar()
        self.tpl_combo = ttk.Combobox(bar, textvariable=self.tpl_var, state="readonly", width=14)
        self.tpl_combo.pack(side="left", padx=2)
        self.tpl_combo.bind("<<ComboboxSelected>>", self._on_tpl_selected)
        ttk.Button(bar, text="模板管理…", command=self.open_template_manager).pack(side="left", padx=2)
        ttk.Button(bar, text="字段映射说明", command=self.show_mapping_doc).pack(side="left", padx=2)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        self.btn_run = ttk.Button(bar, text="开始识别", command=self.start_recognition)
        self.btn_run.pack(side="left", padx=2)
        self.btn_export = ttk.Button(bar, text="导出 Excel…", command=self.export_excel, state="disabled")
        self.btn_export.pack(side="left", padx=2)

        panes = ttk.PanedWindow(self, orient="horizontal")
        panes.pack(fill="both", expand=True, padx=8, pady=(0, 4))

        left = ttk.LabelFrame(panes, text=" 发票文件 ")
        panes.add(left, weight=1)
        self.file_tree = ttk.Treeview(left, columns=("file", "status", "tip"), show="headings")
        for cid, text, w in (("file", "文件名", 280), ("status", "状态", 90), ("tip", "提示", 220)):
            self.file_tree.heading(cid, text=text)
            self.file_tree.column(cid, width=w, anchor="w")
        fsb = ttk.Scrollbar(left, orient="vertical", command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=fsb.set)
        self.file_tree.pack(side="left", fill="both", expand=True)
        fsb.pack(side="right", fill="y")
        self.file_tree.tag_configure("failed", foreground="#c0392b")
        self.file_tree.tag_configure("unsupported", foreground="#c0392b")
        self.file_tree.tag_configure("partial", foreground="#b9770e")
        self.file_tree.tag_configure("ok", foreground="#1e8449")

        right = ttk.LabelFrame(panes, text=" 填表预览（双击单元格可人工修正） ")
        panes.add(right, weight=3)
        self.preview = ttk.Treeview(right, show="headings")
        psb = ttk.Scrollbar(right, orient="vertical", command=self.preview.yview)
        hsb = ttk.Scrollbar(right, orient="horizontal", command=self.preview.xview)
        self.preview.configure(yscrollcommand=psb.set, xscrollcommand=hsb.set)
        self.preview.grid(row=0, column=0, sticky="nsew")
        psb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)
        self.preview.tag_configure("failed", foreground="#c0392b")
        self.preview.tag_configure("unsupported", foreground="#c0392b")
        self.preview.tag_configure("partial", foreground="#b9770e")
        self.preview.bind("<Double-Button-1>", self.on_preview_double_click)

        if _DND:
            left.configure(text=" 发票文件（可直接把文件/文件夹拖到窗口里） ")

        bottom = ttk.Frame(self, padding=(8, 2))
        bottom.pack(fill="x")
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.status_var = tk.StringVar(value="请导入发票文件（PDF / 图片 / 扫描件）")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="right")

        self._rebuild_preview_columns()

    # ---------------- 模板 ----------------
    def reload_templates(self):
        self.templates = load_templates()
        names = [t.name for t in self.templates]
        self.tpl_combo["values"] = names
        if self.current_template is None or self.current_template.id not in [t.id for t in self.templates]:
            self.set_current_template(self.templates[0] if self.templates else None, refresh=False)
        else:
            cur = next(t for t in self.templates if t.id == self.current_template.id)
            self.current_template = cur
            self.tpl_var.set(cur.name)
        self._rebuild_preview_columns()

    def set_current_template(self, tpl: TemplateConfig | None, refresh: bool = True):
        self.current_template = tpl
        if tpl:
            self.tpl_var.set(tpl.name)
        if refresh:
            self._rebuild_preview_columns()

    def _on_tpl_selected(self, _e=None):
        idx = self.tpl_combo.current()
        if 0 <= idx < len(self.templates):
            self.current_template = self.templates[idx]
            self._rebuild_preview_columns()

    # ---------------- 文件 ----------------
    def add_files_dialog(self):
        paths_ = filedialog.askopenfilenames(
            title="选择发票文件（可多选）",
            filetypes=[("发票文件", "*.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff"),
                       ("所有文件", "*.*")])
        self.add_files(list(paths_))

    def add_folder_dialog(self):
        d = filedialog.askdirectory(title="选择发票所在文件夹（自动扫描全部支持的文件）")
        if not d:
            return
        self.add_files(scan_folder(d))

    def add_files(self, paths_: list[str]) -> int:
        added = 0
        for p in paths_:
            if not p.lower().endswith(SUPPORTED_EXT):
                continue
            if p in self.files:
                continue
            self.files.append(p)
            added += 1
        self.refresh_file_tree()
        self.status_var.set(f"已导入 {added} 个新文件，共 {len(self.files)} 个。点击「开始识别」。")
        return added

    # ---------------- 拖拽导入 ----------------
    def _dnd_log(self, msg: str):
        """拖放链路日志，便于排查"拖了没反应"类问题。"""
        try:
            d = paths.logs_dir()
            os.makedirs(d, exist_ok=True)
            with open(os.path.join(d, "dnd.log"), "a", encoding="utf-8") as f:
                import datetime
                f.write(f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")
        except OSError:
            pass

    def _enable_drag_drop(self):
        """把窗口和所有子控件都注册为拖放目标：从资源管理器拖文件/文件夹到窗口任意位置即可。

        注意：<<Drop*>> 事件必须用 tkinterdnd2 的 dnd_bind（提供 %D 数据替换），
        用普通 bind 拿不到 event.data，处理函数会静默失败。
        """
        def register(w):
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<DropEnter>>", self._on_drop_enter)
            w.dnd_bind("<<DropLeave>>", self._on_drop_leave)
            w.dnd_bind("<<Drop>>", self._on_drop)
            for child in w.winfo_children():
                register(child)
        register(self)
        ver = getattr(self, "TkdndVersion", "?")
        self._dnd_log(f"拖放已启用（tkdnd {ver}）。管理员身份运行会禁用资源管理器拖放。")

    def _on_drop_enter(self, event):
        self.status_var.set("检测到拖入，松开鼠标即可导入文件/文件夹…")
        return getattr(event, "action", None) or "copy"

    def _on_drop_leave(self, event):
        self.status_var.set("已取消拖入。")
        return "copy"

    def _parse_drop_data(self, data) -> list[str]:
        if isinstance(data, (tuple, list)):
            return [str(x) for x in data]
        text = str(data).strip()
        if not text:
            return []
        try:
            return [str(x) for x in self.tk.splitlist(text)]
        except Exception:
            import re
            out = []
            for m in re.finditer(r"\{([^}]*)\}|(\S+)", text):
                out.append(m.group(1) if m.group(1) is not None else m.group(2))
            return out

    def _on_drop(self, event):
        try:
            items = self._parse_drop_data(event.data)
            self._dnd_log(f"收到拖放：{items}")
            files: list[str] = []
            folders = 0
            for item in items:
                p = os.path.normpath(str(item))
                if os.path.isdir(p):
                    folders += 1
                    files.extend(scan_folder(p))
                elif os.path.isfile(p) and p.lower().endswith(SUPPORTED_EXT):
                    files.append(p)
            if not files:
                self.status_var.set(
                    "拖入的文件夹里没有可识别的发票文件" if folders
                    else "拖入的内容不是支持的类型（PDF / 图片 / 扫描件）")
                self._dnd_log("无可识别文件")
                return "copy"
            added = self.add_files(files)
            src = "文件夹" if folders else "文件"
            self.status_var.set(
                f"拖入{src}成功：新增 {added} 个（共 {len(self.files)} 个），点击「开始识别」。")
        except Exception as e:  # 拖放回调异常不能无声吞掉
            self._dnd_log(f"拖放处理异常：{type(e).__name__}: {e}")
            self.status_var.set(f"拖入处理失败：{e}")
        return "copy"

    def remove_selected_files(self):
        sel = {self.file_tree.index(i) for i in self.file_tree.selection()}
        if not sel:
            return
        kept = [(i, f) for i, f in enumerate(self.files) if i not in sel]
        self.files = [f for _, f in kept]
        self.records = {new: self.records[old]
                        for new, (old, _f) in enumerate(kept) if old in self.records}
        self.refresh_file_tree()
        self.rebuild_preview_rows()

    def refresh_file_tree(self):
        self.file_tree.delete(*self.file_tree.get_children())
        for i, f in enumerate(self.files):
            rec = self.records.get(i)
            if rec is None:
                vals = (os.path.basename(f), "待识别", "")
                tag = ()
            else:
                tip = "；".join(rec.warnings)
                vals = (os.path.basename(f), STATUS_TEXT.get(rec.status, rec.status), tip)
                tag = (rec.status,)
            self.file_tree.insert("", "end", iid=str(i), values=vals, tags=tag)

    # ---------------- 识别 ----------------
    def start_recognition(self):
        if self.running:
            return
        if not self.files:
            messagebox.showinfo("提示", "请先导入发票文件")
            return
        self.running = True
        self.btn_run["state"] = "disabled"
        self.progress.configure(maximum=len(self.files), value=0)
        self.status_var.set("识别中…（PDF 文字层优先，扫描件/图片走本地 OCR）")
        threading.Thread(target=self._worker, args=(list(enumerate(self.files)),),
                         daemon=True).start()

    def _worker(self, jobs: list[tuple[int, str]]):
        for idx, path in jobs:
            rec = parse_invoice(path)
            self.queue.put(("file", idx, rec))
        self.queue.put(("done",))

    def _poll_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg[0] == "file":
                    _, idx, rec = msg
                    self.records[idx] = rec
                    rec_item = self.file_tree.exists(str(idx))
                    if rec_item:
                        self.refresh_file_tree()
                    self.update_preview_row(idx)
                    self.progress.step(1)
                elif msg[0] == "done":
                    self.running = False
                    self.btn_run["state"] = "normal"
                    ok = sum(1 for r in self.records.values() if r.status == "ok")
                    self.status_var.set(
                        f"识别完成：{ok}/{len(self.files)} 成功。可双击预览单元格修正，然后导出 Excel。")
                    self.btn_export["state"] = "normal" if self.records else "disabled"
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    # ---------------- 预览 ----------------
    def _rebuild_preview_columns(self):
        cols = ["源文件"]
        if self.current_template:
            cols += [c.header for c in self.current_template.columns]
        cols += ["状态"]
        self.preview["columns"] = cols
        for cid in cols:
            self.preview.heading(cid, text=cid)
            self.preview.column(cid, width=150, anchor="w", stretch=(cid != "源文件"))
        self.preview.column("源文件", width=200, stretch=False)
        self.rebuild_preview_rows()

    def rebuild_preview_rows(self):
        self.preview.delete(*self.preview.get_children())
        for i in range(len(self.files)):
            self.update_preview_row(i, insert_only=True)

    def _row_values(self, i: int, rec: InvoiceData | None) -> tuple:
        vals = [os.path.basename(self.files[i])]
        if rec is None:
            vals += [""] * (len(self.current_template.columns) if self.current_template else 0)
            vals += ["待识别"]
        else:
            for col in (self.current_template.columns if self.current_template else []):
                if col.field == "_row_number":
                    vals.append(str(i + 1))
                elif col.field:
                    vals.append(rec.get_field(col.field))
                else:
                    vals.append("")
            vals.append(STATUS_TEXT.get(rec.status, rec.status))
        return tuple(vals)

    def update_preview_row(self, i: int, insert_only: bool = False):
        if not self.current_template:
            return
        iid = str(i)
        vals = self._row_values(i, self.records.get(i))
        rec = self.records.get(i)
        tag = (rec.status,) if rec else ()
        if self.preview.exists(iid) and not insert_only:
            self.preview.item(iid, values=vals, tags=tag)
        elif not self.preview.exists(iid):
            self.preview.insert("", "end", iid=iid, values=vals, tags=tag)

    def on_preview_double_click(self, event):
        if not self.current_template or not self.preview.selection():
            return
        iid = self.preview.selection()[0]
        col_id = self.preview.identify_column(event.x)
        col_n = int(col_id.replace("#", "") or 0) - 1
        headers = ["源文件"] + [c.header for c in self.current_template.columns] + ["状态"]
        if col_n < 1 or col_n >= len(headers) - 1:
            return
        col = self.current_template.columns[col_n - 1]
        idx = int(iid)
        rec = self.records.get(idx)
        if rec is None:
            messagebox.showinfo("提示", "该文件尚未识别")
            return
        if not col.field or col.field == "_row_number":
            messagebox.showinfo("提示", f"「{col.header}」为自动生成或无映射列，不可修改")
            return
        current = rec.get_field(col.field)
        EditDialog(self, iid, col.field, col.header, current)

    def commit_edit(self, iid: str, field_key: str, value: str):
        idx = int(iid)
        rec = self.records.get(idx)
        if rec is None:
            return
        rec.set_field(field_key, value)
        if field_key in rec.overrides and "人工已修正" not in rec.warnings:
            rec.warnings.append("人工已修正")
        self.update_preview_row(idx)
        self.refresh_file_tree()
        self.status_var.set("已修正（导出时以修正值为准）")

    # ---------------- 导出 ----------------
    def export_excel(self):
        if not self.records:
            messagebox.showinfo("提示", "请先识别发票")
            return
        tpl = self.current_template
        if not tpl:
            return
        default_name = f"填表结果.xlsx"
        out = filedialog.asksaveasfilename(
            title="导出填表结果", defaultextension=".xlsx", initialfile=default_name,
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not out:
            return
        records = [self.records.get(i) for i in range(len(self.files))]
        records = [r for r in records if r is not None]
        try:
            stats = export(records, tpl, out)
        except PermissionError:
            messagebox.showerror("导出失败", "目标 Excel 文件正被打开，请关闭后重试。")
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("导出失败", f"{type(e).__name__}: {e}")
            return
        md = generate_mapping_doc(tpl, records)
        md_path = os.path.splitext(out)[0] + "_字段映射说明.md"
        try:
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(md)
        except OSError:
            md_path = ""
        msg = (f"已导出：{out}\n\n成功 {stats.get('ok', 0)} / 部分 {stats.get('partial', 0)} / "
               f"失败 {stats.get('failed', 0)} / 版式不支持 {stats.get('unsupported', 0)}\n"
               "缺失字段已留白并在「处理报告」工作表中标注。")
        if md_path:
            msg += f"\n字段映射说明：{md_path}"
        if messagebox.askyesno("导出完成", msg + "\n\n打开所在文件夹？"):
            try:
                os.startfile(os.path.dirname(os.path.abspath(out)))  # noqa: S606
            except OSError:
                pass

    # ---------------- 映射说明 / 模板管理 ----------------
    def show_mapping_doc(self):
        tpl = self.current_template
        if not tpl:
            return
        md = generate_mapping_doc(tpl, list(self.records.values()))
        win = tk.Toplevel(self)
        win.title(f"字段映射说明 - {tpl.name}")
        win.geometry("860x560")
        txt = tk.Text(win, wrap="none", font=("Microsoft YaHei UI", 10))
        ysb = ttk.Scrollbar(win, orient="vertical", command=txt.yview)
        xsb = ttk.Scrollbar(win, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        txt.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        win.rowconfigure(0, weight=1)
        win.columnconfigure(0, weight=1)
        txt.insert("1.0", md)
        txt.configure(state="disabled")

    def open_template_manager(self):
        TemplateManager(self)


def main():
    paths.ensure_builtin_template()
    app = InvoiceFillerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
