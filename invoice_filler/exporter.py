# -*- coding: utf-8 -*-
"""按模板导出 Excel：保留原模板表头顺序/样式/合并单元格，缺失字段留白并标注，
另加「处理报告」工作表汇总每份发票的识别状态，不影响模板原有工作表。"""
from __future__ import annotations

import datetime
import os
import re
from typing import List, Optional, Tuple

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string

from .fields import ALL_FIELDS, NUMBER_FORMATS
from .models import STATUS_TEXT, InvoiceData
from .template import TemplateConfig

REPORT_SHEET = "处理报告"
AUTHOR = "发票填表工具"


def _parse_money(s: str) -> Optional[float]:
    t = (s or "").replace("¥", "").replace("￥", "").replace(",", "").strip()
    try:
        return float(t)
    except ValueError:
        return None


def _parse_date(s: str):
    t = (s or "").strip()
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", t)
    if not m:
        m2 = re.match(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", t)
        m = m2
    if not m:
        return None
    try:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def convert_value(raw: str, value_type: str) -> Tuple[object, bool]:
    """把字段文本转换为写入值；返回 (值, 是否成功)。空文本返回 (None, True)。"""
    raw = (raw or "").strip()
    if raw == "":
        return None, True
    if value_type == "money":
        v = _parse_money(raw)
        return (v, True) if v is not None else (None, False)
    if value_type == "date":
        v = _parse_date(raw)
        return (v, True) if v is not None else (None, False)
    if value_type == "int":
        try:
            return int(float(raw)), True
        except ValueError:
            return None, False
    return raw, True


def _base_font(ws, col_index: int) -> Font:
    """数据行字体沿用表头字体（去掉加粗）。"""
    h = ws.cell(row=1, column=col_index)
    f = h.font
    return Font(name=f.name, size=f.size, bold=False, color=f.color)


def _header_border(ws, col_index: int):
    return ws.cell(row=1, column=col_index).border.copy() if ws.cell(
        row=1, column=col_index).border else None


def build_workbook_from_template(template: TemplateConfig) -> "openpyxl.Workbook":
    """模板源文件丢失时，按模板 JSON 重建一份同样式的空白工作簿。"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = template.sheet
    for c in template.columns:
        cell = ws.cell(row=template.header_row, column=c.col_index, value=c.header)
        cell.font = Font(name="宋体", size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if c.colspan > 1:
            ws.merge_cells(start_row=template.header_row, start_column=c.col_index,
                           end_row=template.header_row, end_column=c.col_index + c.colspan - 1)
    ws.row_dimensions[template.header_row].height = 30
    return wb


def export(records: List[InvoiceData], template: TemplateConfig, out_path: str) -> dict:
    src = template.source_file
    if src and os.path.exists(src):
        wb = openpyxl.load_workbook(src)
    else:
        wb = build_workbook_from_template(template)
    if template.sheet not in wb.sheetnames:
        wb.create_sheet(template.sheet)
    ws = wb[template.sheet]

    data_start = template.data_start_row
    # 防覆盖：若目标工作表在起始行之后已有数据（如直接写入月度台账副本），
    # 自动顺延到第一个空行，绝不覆盖既有记录。
    def _row_empty(r: int) -> bool:
        return all(ws.cell(row=r, column=c.col_index).value in (None, "")
                   for c in template.columns)
    r = data_start
    while r <= ws.max_row and not _row_empty(r):
        r += 1
    data_start = r
    stats = {"ok": 0, "partial": 0, "failed": 0, "unsupported": 0}
    missing_log: List[List[str]] = []

    for i, rec in enumerate(records):
        r = data_start + i
        stats[rec.status] = stats.get(rec.status, 0) + 1
        row_missing: List[str] = []

        for col in template.columns:
            cell = ws.cell(row=r, column=col.col_index)
            fdef = ALL_FIELDS.get(col.field) if col.field else None
            value_type = fdef["value_type"] if fdef else "text"
            label = fdef["label"] if fdef else col.header

            if col.field == "_row_number":
                raw = str(i + 1)
            elif col.field:
                raw = rec.get_field(col.field)
            else:
                raw = ""      # 无映射 → 留白

            value, ok = convert_value(raw, value_type)
            if raw and not ok:
                row_missing.append(f"{col.header}（格式异常：{raw}）")
                continue
            if value is None:
                if col.field and col.field != "_row_number":
                    # 期望有值但缺失 → 留白并批注标注
                    note = "未识别到「%s」，已留白" % label
                    if rec.status == "unsupported":
                        note = "版式不支持，已留白"
                    cell.comment = Comment(note, AUTHOR)
                    row_missing.append(label)
                continue

            cell.value = value
            nf = NUMBER_FORMATS.get(value_type)
            if nf:
                cell.number_format = nf
            cell.font = _base_font(ws, col.col_index)
            border = _header_border(ws, col.col_index)
            if border:
                cell.border = border
            if value_type in ("text", "rate", "int"):
                cell.alignment = Alignment(horizontal="center" if value_type == "int" else "left",
                                           vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(vertical="center")
            if col.colspan > 1:
                ws.merge_cells(start_row=r, start_column=col.col_index,
                               end_row=r, end_column=col.col_index + col.colspan - 1)
        if row_missing:
            missing_log.append([os.path.basename(rec.source_file), row_missing])

    # ---- 处理报告表 ----
    if REPORT_SHEET in wb.sheetnames:
        del wb[REPORT_SHEET]
    rep = wb.create_sheet(REPORT_SHEET)
    headers = ["文件名", "处理状态", "缺失/留白字段", "警告信息"]
    for j, h in enumerate(headers, start=1):
        c = rep.cell(row=1, column=j, value=h)
        c.font = Font(name="宋体", size=11, bold=True)
        c.fill = PatternFill("solid", fgColor="F2F2F2")
        c.alignment = Alignment(horizontal="center", vertical="center")
    rep.row_dimensions[1].height = 22
    for j, w in enumerate([42, 14, 32, 60], start=1):
        rep.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    for i, rec in enumerate(records):
        missing = next((m[1] for m in missing_log if m[0] == os.path.basename(rec.source_file)), [])
        rep.cell(row=2 + i, column=1, value=os.path.basename(rec.source_file))
        rep.cell(row=2 + i, column=2, value=STATUS_TEXT.get(rec.status, rec.status))
        rep.cell(row=2 + i, column=3, value="、".join(missing))
        rep.cell(row=2 + i, column=4, value="；".join(rec.warnings))
        for j in range(1, 5):
            rep.cell(row=2 + i, column=j).font = Font(name="宋体", size=10)
            rep.cell(row=2 + i, column=j).alignment = Alignment(vertical="center", wrap_text=True)
    rep.freeze_panes = "A2"

    wb.properties.creator = AUTHOR
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)
    stats["total"] = len(records)
    stats["out_path"] = out_path
    return stats
