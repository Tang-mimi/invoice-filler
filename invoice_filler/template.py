# -*- coding: utf-8 -*-
"""Excel 模板映射配置：加载/保存 JSON、从 xlsx 表头自动生成映射。

模板 JSON 结构（模板1 为内置示例）：
{
  "id": "template1", "name": "模板1",
  "source_file": "填表模板.xlsx", "sheet": "Sheet1",
  "header_row": 1, "data_start_row": 2,
  "columns": [ {"header":..., "col":"B", "col_index":2, "colspan":1,
                "field":"issue_date"| null, "note":"..."} , ...]
}
新增模板只需导入新的 xlsx 并确认映射，生成新的 JSON，主流程零改动。
"""
from __future__ import annotations

import json
import os
import uuid
from dataclasses import dataclass, field as dc_field
from typing import Dict, List, Optional

from . import paths
from .fields import ALL_FIELDS, match_header, normalize_header

TEMPLATE_VERSION = 1


@dataclass
class ColumnSpec:
    header: str
    col: str                 # 列字母，如 "B"
    col_index: int           # 1-based 列号
    colspan: int = 1         # 表头横向合并的列数
    field: Optional[str] = None   # 规范字段 key；None=无映射（留白）
    note: str = ""           # 映射说明（自动生成，可编辑）

    def to_dict(self) -> dict:
        return {"header": self.header, "col": self.col, "col_index": self.col_index,
                "colspan": self.colspan, "field": self.field, "note": self.note}

    @staticmethod
    def from_dict(d: dict) -> "ColumnSpec":
        return ColumnSpec(header=d.get("header", ""), col=d.get("col", ""),
                          col_index=int(d.get("col_index", 1)),
                          colspan=int(d.get("colspan", 1)),
                          field=d.get("field"), note=d.get("note", ""))


@dataclass
class TemplateConfig:
    id: str
    name: str
    source_file: str = ""
    sheet: str = "Sheet1"
    header_row: int = 1
    data_start_row: int = 2
    columns: List[ColumnSpec] = dc_field(default_factory=list)
    version: int = TEMPLATE_VERSION

    # ---------- 序列化 ----------
    def to_dict(self) -> dict:
        return {"id": self.id, "name": self.name, "source_file": self.source_file,
                "sheet": self.sheet, "header_row": self.header_row,
                "data_start_row": self.data_start_row, "version": self.version,
                "columns": [c.to_dict() for c in self.columns]}

    def save(self) -> str:
        path = os.path.join(paths.templates_dir(), f"{self.id}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)
        return path

    @staticmethod
    def from_dict(d: dict) -> "TemplateConfig":
        return TemplateConfig(
            id=d.get("id", f"tpl_{uuid.uuid4().hex[:8]}"),
            name=d.get("name", "未命名模板"),
            source_file=d.get("source_file", ""),
            sheet=d.get("sheet", "Sheet1"),
            header_row=int(d.get("header_row", 1)),
            data_start_row=int(d.get("data_start_row", 2)),
            columns=[ColumnSpec.from_dict(c) for c in d.get("columns", [])],
            version=int(d.get("version", 1)),
        )

    @staticmethod
    def load(path: str) -> "TemplateConfig":
        with open(path, "r", encoding="utf-8") as f:
            return TemplateConfig.from_dict(json.load(f))

    def field_label(self, key: Optional[str]) -> str:
        if not key:
            return "（无映射，留白）"
        f = ALL_FIELDS.get(key)
        return f["label"] if f else key


def load_templates() -> List[TemplateConfig]:
    """读取模板目录下全部 JSON（内置模板1 首次运行自动释放）。"""
    paths.ensure_builtin_template()
    out: List[TemplateConfig] = []
    d = paths.templates_dir()
    for fn in sorted(os.listdir(d)):
        if fn.endswith(".json"):
            try:
                out.append(TemplateConfig.load(os.path.join(d, fn)))
            except Exception:  # 坏文件跳过，不影响其他模板
                continue
    out.sort(key=lambda t: (t.id != "template1", t.name))
    return out


def build_template_from_workbook(path: str, sheet: Optional[str] = None,
                                 header_row: int = 1, name: str = "") -> TemplateConfig:
    """从用户提供的 xlsx 读取表头，自动匹配字段，生成新模板。"""
    import openpyxl
    from openpyxl.utils import get_column_letter, column_index_from_string

    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
        sheet = ws.title
        merged: Dict[int, int] = {}   # 起始列号 -> 跨度
        merged_starts = set()
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= header_row <= rng.max_row and rng.min_row == rng.max_row:
                merged[rng.min_col] = rng.max_col - rng.min_col + 1
                for c in range(rng.min_col + 1, rng.max_col + 1):
                    merged_starts.add(c)

        columns: List[ColumnSpec] = []
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=header_row, column=c)
            if c in merged_starts:
                continue  # 合并区域的非起始列
            header = str(cell.value).strip() if cell.value is not None else ""
            if not header:
                continue
            span = merged.get(c, 1)
            fld = match_header(header)
            columns.append(ColumnSpec(
                header=header, col=get_column_letter(c), col_index=c, colspan=span,
                field=fld["key"] if fld else None,
                note=(fld["source"] if fld else "发票中无对应字段，留白")))
        tpl = TemplateConfig(
            id=f"tpl_{uuid.uuid4().hex[:8]}",
            name=name or f"模板-{os.path.splitext(os.path.basename(path))[0]}",
            source_file=os.path.abspath(path), sheet=sheet, header_row=header_row,
            data_start_row=header_row + 1, columns=columns)
        return tpl
    finally:
        wb.close()


def reassign_field(template: TemplateConfig, header: str, field_key: Optional[str]) -> str:
    """人工调整某表头的映射，返回说明文字。"""
    if field_key:
        return ALL_FIELDS[field_key]["source"]
    return "发票中无对应字段，留白"


def match_quality(template: TemplateConfig) -> Dict[str, int]:
    """统计映射情况，用于导入模板后的确认提示。"""
    mapped = sum(1 for c in template.columns if c.field)
    return {"total": len(template.columns), "mapped": mapped,
            "unmapped": len(template.columns) - mapped}
