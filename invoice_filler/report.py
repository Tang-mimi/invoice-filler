# -*- coding: utf-8 -*-
"""字段映射说明生成：供用户确认每个表头与发票原始内容的对应关系。"""
from __future__ import annotations

import datetime
from typing import List, Optional

from .fields import ALL_FIELDS
from .models import InvoiceData
from .template import TemplateConfig


def _example_value(records: List[InvoiceData], field_key: Optional[str]) -> str:
    if not field_key:
        return "—"
    for rec in records:
        v = rec.get_field(field_key)
        if v:
            return v.replace("\n", " ")
    return "（本批样例无值）"


def generate_mapping_doc(template: TemplateConfig,
                         records: Optional[List[InvoiceData]] = None) -> str:
    records = records or []
    lines = []
    lines.append(f"# 字段映射说明（{template.name}）\n")
    lines.append(f"- 生成时间：{datetime.datetime.now():%Y-%m-%d %H:%M}")
    lines.append(f"- 模板来源：{template.source_file or '（按模板配置重建）'}"
                 f" / 工作表「{template.sheet}」"
                 f" / 表头第{template.header_row}行，数据从第{template.data_start_row}行开始\n")
    lines.append("| Excel列 | 表头 | 映射字段 | 发票中的原始内容 | 本批示例值 | 状态 |")
    lines.append("|---|---|---|---|---|---|")
    for col in template.columns:
        if col.field:
            f = ALL_FIELDS.get(col.field, {})
            label = f.get("label", col.field)
            source = col.note or f.get("source", "")
            status = "已映射"
        else:
            label, source, status = "（无映射）", "发票中无对应原始内容", "留白"
        col_span = f"{col.col}" if col.colspan <= 1 else f"{col.col}~{chr(ord(col.col)+col.colspan-1)}"
        example = _example_value(records, col.field)
        if col.field == "_row_number":
            example = "1（自动编号）"
        lines.append(f"| {col_span} | {col.header} | {label} | {source} | {example} | {status} |")
    lines.append("")
    lines.append("## 说明")
    lines.append("1. 「客户信息/客户名称」列固定取发票「购买方信息-名称」。")
    lines.append("2. 除自动编号外，每个已映射表头均对应发票上明确的原始内容；"
                 "无映射的表头（如本模板的「交易模式」在发票上没有对应栏目）一律留白。")
    lines.append("3. 数电票（电子发票）没有「发票代码」，涉及该字段的列将留空。")
    lines.append("4. 一张发票含多行明细时，明细类字段（品名、规格、数量等）按行以「；」连接。")
    lines.append("5. 识别失败或缺失的字段留白，并在单元格批注及「处理报告」表中标注；"
                 "预览界面双击可人工修正，修正值优先于识别值。")
    lines.append("6. 金额、税额、价税合计为数值（元），开票日期为 yyyy-mm-dd 日期格式，"
                 "发票号码按文本写入防止位数丢失。")
    return "\n".join(lines)
