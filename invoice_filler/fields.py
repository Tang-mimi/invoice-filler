# -*- coding: utf-8 -*-
"""发票规范字段注册表与表头自动匹配。

新增模板的表头通过与这里的别名表自动匹配生成映射；匹配不上的列留白，
可在模板管理界面人工指定。扩展新字段只需在 FIELD_REGISTRY/PSEUDO_FIELDS
中加条目，主流程代码无需改动。
"""
from __future__ import annotations

import re
from typing import Optional

# value_type: text=按文本写入(发票号等加'@'格式)  money=数值(千分位两位小数)
#             date=日期(yyyy-mm-dd)  int=整数  rate=文本(保留发票原样如"13%")
FIELD_REGISTRY = [
    {"key": "invoice_number", "label": "发票号码", "value_type": "text",
     "aliases": ["发票号码", "发票号", "数电票号码", "号码"],
     "source": "发票右上角「发票号码」（数电票为20位数字）"},
    {"key": "invoice_code", "label": "发票代码", "value_type": "text",
     "aliases": ["发票代码", "代码"],
     "source": "老版发票左上角「发票代码」；数电票无此字段，留空"},
    {"key": "issue_date", "label": "开票日期", "value_type": "date",
     "aliases": ["开票日期", "填开日期", "开票时间", "日期"],
     "source": "发票右上角「开票日期」"},
    {"key": "buyer_name", "label": "购买方名称", "value_type": "text",
     "aliases": ["客户信息", "客户名称", "客户", "客户全称", "购买方名称", "购买方信息-名称",
                 "购方名称", "购方", "购买方", "购方单位名称"],
     "source": "「购买方信息」栏 -「名称」"},
    {"key": "buyer_tax_id", "label": "购买方纳税人识别号", "value_type": "text",
     "aliases": ["纳税人识别号", "购买方纳税人识别号", "购方税号", "购买方税号", "客户税号",
                 "税号", "统一社会信用代码", "购买方统一社会信用代码"],
     "source": "「购买方信息」栏 -「统一社会信用代码/纳税人识别号」"},
    {"key": "seller_name", "label": "销售方名称", "value_type": "text",
     "aliases": ["销售方名称", "销售方信息-名称", "销方名称", "销售方", "销方",
                 "供应商名称", "供应商"],
     "source": "「销售方信息」栏 -「名称」"},
    {"key": "seller_tax_id", "label": "销售方纳税人识别号", "value_type": "text",
     "aliases": ["销售方纳税人识别号", "销方税号", "销售方税号", "供应商税号",
                 "销售方统一社会信用代码"],
     "source": "「销售方信息」栏 -「统一社会信用代码/纳税人识别号」"},
    {"key": "item_names", "label": "货物或服务名称", "value_type": "text",
     "aliases": ["货物或服务名称", "项目名称", "货物名称", "商品名称", "服务名称", "品名",
                 "货物或应税劳务、服务名称", "内容"],
     "source": "货物明细「项目名称」，多行明细以「；」连接"},
    {"key": "item_specs", "label": "规格型号", "value_type": "text",
     "aliases": ["规格型号", "规格", "型号"],
     "source": "货物明细「规格型号」，多行明细以「；」连接"},
    {"key": "item_units", "label": "单位", "value_type": "text",
     "aliases": ["单位", "计量单位"], "source": "货物明细「单位」"},
    {"key": "item_qtys", "label": "数量", "value_type": "text",
     "aliases": ["数量"], "source": "货物明细「数量」，多行明细以「；」连接"},
    {"key": "item_unit_prices", "label": "单价", "value_type": "text",
     "aliases": ["单价"], "source": "货物明细「单价」"},
    {"key": "total_amount", "label": "金额（不含税合计）", "value_type": "money",
     "aliases": ["金额", "合计金额", "不含税金额", "金额(不含税)", "金额（不含税）",
                 "合计金额(不含税)", "货物金额"],
     "source": "「合计」行的「金额」（不含税）"},
    {"key": "tax_rate", "label": "税率", "value_type": "rate",
     "aliases": ["税率", "征收率", "税率/征收率", "税率(征收率)"],
     "source": "货物明细「税率/征收率」，多个税率时去重后以「；」连接"},
    {"key": "total_tax", "label": "税额", "value_type": "money",
     "aliases": ["税额", "合计税额"], "source": "「合计」行的「税额」"},
    {"key": "total_price_tax", "label": "价税合计（小写）", "value_type": "money",
     "aliases": ["价税合计", "价税合计（小写）", "价税合计(小写)", "含税金额",
                 "价税合计小写", "合计金额(含税)", "合计金额（含税）"],
     "source": "「价税合计」行的「（小写）」金额"},
    {"key": "total_price_tax_cn", "label": "价税合计（大写）", "value_type": "text",
     "aliases": ["价税合计（大写）", "价税合计(大写)", "大写金额", "价税合计大写"],
     "source": "「价税合计」行的「（大写）」文字"},
    {"key": "remark", "label": "备注", "value_type": "text",
     "aliases": ["备注", "备注栏", "备注信息"], "source": "发票「备注」栏全文"},
    {"key": "issuer", "label": "开票人", "value_type": "text",
     "aliases": ["开票人"], "source": "发票底部「开票人」"},
    {"key": "buyer_bank", "label": "购方开户银行", "value_type": "text",
     "aliases": ["购方开户银行", "购买方开户银行", "购方开户行", "购买方开户行"],
     "source": "「备注」栏 - 购方开户银行"},
    {"key": "buyer_account", "label": "购方银行账号", "value_type": "text",
     "aliases": ["购方银行账号", "购买方银行账号", "购方账号", "购买方账号"],
     "source": "「备注」栏 - 购方银行账号"},
    {"key": "seller_bank", "label": "销方开户银行", "value_type": "text",
     "aliases": ["销方开户银行", "销售方开户银行", "销方开户行", "销售方开户行"],
     "source": "「备注」栏 - 销方开户银行"},
    {"key": "seller_account", "label": "销方银行账号", "value_type": "text",
     "aliases": ["销方银行账号", "销售方银行账号", "销方账号", "销售方账号"],
     "source": "「备注」栏 - 销方银行账号"},
]

# 伪字段：不来自发票内容，导出时自动生成
PSEUDO_FIELDS = [
    {"key": "_row_number", "label": "序号（自动编号）", "value_type": "int",
     "aliases": ["序号", "编号", "序列", "no", "row"],
     "source": "导出时按文件导入顺序自动编号（1、2、3…），非发票内容"},
    {"key": "_source_file", "label": "源文件名", "value_type": "text",
     "aliases": ["文件名", "源文件", "源文件名", "发票文件", "附件"],
     "source": "导入的发票文件文件名"},
]

ALL_FIELDS = {f["key"]: f for f in FIELD_REGISTRY + PSEUDO_FIELDS}

# value_type -> 写入 Excel 时的 number_format；None 表示不设置
NUMBER_FORMATS = {"text": "@", "money": "#,##0.00", "date": "yyyy-mm-dd",
                  "int": "0", "rate": "@"}

# 目前支持解析的发票版式（标题关键词）；不在范围内的标注"版式不支持"并留空
SUPPORTED_TITLE_KEYWORDS = ["电子发票", "数电", "全电"]


def normalize_header(text: str) -> str:
    """表头归一化：去空白、全角括号冒号转半角、小写。"""
    if text is None:
        return ""
    t = str(text)
    t = re.sub(r"\s+", "", t)
    t = t.replace("（", "(").replace("）", ")").replace("：", ":").replace("／", "/")
    return t.lower()


def match_header(header: str) -> Optional[dict]:
    """把一个 Excel 表头匹配到规范字段。返回字段定义 dict，无匹配返回 None。

    规则：别名完全相等 > 别名与表头互为包含（取最长别名）。
    """
    norm = normalize_header(header)
    if not norm:
        return None
    best = None
    best_score = 0.0
    for f in FIELD_REGISTRY + PSEUDO_FIELDS:
        for alias in f["aliases"]:
            a = normalize_header(alias)
            if norm == a:
                score = 1.0
            elif len(a) >= 2 and (a in norm or norm in a):
                score = 0.4 + 0.1 * min(len(a), len(norm)) / 10
            else:
                continue
            if score > best_score:
                best, best_score = f, score
    return best
