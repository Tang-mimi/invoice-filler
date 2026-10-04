# -*- coding: utf-8 -*-
"""识别结果数据模型。"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class Item:
    """发票货物/服务明细行。数值保留发票原样文本，导出时按需转换。"""
    name: str = ""
    spec: str = ""
    unit: str = ""
    qty: str = ""
    price: str = ""
    amount: str = ""
    tax_rate: str = ""
    tax_amount: str = ""


# 文件级处理状态
ST_PENDING = "pending"
ST_OK = "ok"                # 全部核心字段识别成功
ST_PARTIAL = "partial"      # 识别成功但部分字段缺失
ST_FAILED = "failed"        # 识别失败（异常/未检测到发票内容）
ST_UNSUPPORTED = "unsupported"  # 版式不支持

STATUS_TEXT = {
    ST_PENDING: "待识别",
    ST_OK: "成功",
    ST_PARTIAL: "部分字段缺失",
    ST_FAILED: "识别失败",
    ST_UNSUPPORTED: "版式不支持",
}

# 判定"识别基本成功"所需的核心字段
CORE_FIELDS = ["invoice_number", "issue_date", "buyer_name", "total_price_tax"]


@dataclass
class InvoiceData:
    source_file: str = ""
    invoice_type: str = ""          # 发票标题，如 电子发票（增值税专用发票）
    invoice_number: str = ""
    invoice_code: str = ""
    issue_date: str = ""            # 归一化为 yyyy-mm-dd
    buyer_name: str = ""
    buyer_tax_id: str = ""
    seller_name: str = ""
    seller_tax_id: str = ""
    items: List[Item] = field(default_factory=list)
    total_amount: str = ""
    total_tax: str = ""
    total_price_tax: str = ""
    total_price_tax_cn: str = ""
    tax_rate: str = ""
    remark: str = ""
    issuer: str = ""
    buyer_bank: str = ""
    buyer_account: str = ""
    seller_bank: str = ""
    seller_account: str = ""
    status: str = ST_PENDING
    warnings: List[str] = field(default_factory=list)
    overrides: dict = field(default_factory=dict)   # 人工修正值，优先于识别值
    engine: str = ""                # text-layer / ocr

    # ---------- 字段读写 ----------
    def get_field(self, key: str) -> str:
        """按规范字段 key 取值；人工修正优先。"""
        if key in self.overrides:
            return self.overrides[key]
        if key == "_source_file":
            return os.path.basename(self.source_file)
        if key == "_row_number":
            return ""  # 导出器按行序填写
        if key.startswith("item_"):
            attr = {"item_names": "name", "item_specs": "spec", "item_units": "unit",
                    "item_qtys": "qty", "item_unit_prices": "price",
                    "item_amounts": "amount", "item_tax_rates": "tax_rate",
                    "item_tax_amounts": "tax_amount"}.get(key)
            if attr is None:
                return ""
            vals = [getattr(it, attr) for it in self.items if getattr(it, attr)]
            return "；".join(vals)
        return str(getattr(self, key, "") or "")

    def set_field(self, key: str, value: str) -> None:
        """人工修正：统一写入 overrides，导出与预览均以修正值为准。"""
        value = (value or "").strip()
        if value:
            self.overrides[key] = value
        else:
            self.overrides.pop(key, None)

    # ---------- 状态维护 ----------
    def finalize_status(self) -> None:
        if self.status in (ST_FAILED, ST_UNSUPPORTED):
            return
        missing = [k for k in CORE_FIELDS if not self.get_field(k)]
        if not self.items and not missing:
            self.warnings.append("未解析出货物明细行（不影响金额合计）")
        if missing:
            labels = {"invoice_number": "发票号码", "issue_date": "开票日期",
                      "buyer_name": "购买方名称", "total_price_tax": "价税合计"}
            self.warnings.append("缺失：" + "、".join(labels[m] for m in missing))
            self.status = ST_PARTIAL
        else:
            self.status = ST_OK

    def missing_core_labels(self) -> List[str]:
        labels = {"invoice_number": "发票号码", "issue_date": "开票日期",
                  "buyer_name": "购买方名称", "total_price_tax": "价税合计"}
        return [labels[k] for k in CORE_FIELDS if not self.get_field(k)]

    def summary(self) -> str:
        n = self.invoice_number or "(无号码)"
        b = self.buyer_name or "(无购方)"
        amt = self.total_price_tax or "(无价税合计)"
        return f"{n} | {b} | 价税合计 {amt}"
