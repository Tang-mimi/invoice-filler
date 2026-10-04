# -*- coding: utf-8 -*-
"""命令行批处理/校验入口（也是开发验证工具）。

用法：
  python tools/cli_check.py 发票1.pdf 发票2.png ... [--template templates/template1.json]
                            [--out 输出/结果.xlsx] [--dump]
不带 --out 时仅打印识别结果，不写 Excel。
"""
from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from invoice_filler.exporter import export            # noqa: E402
from invoice_filler.parser import parse_invoice       # noqa: E402
from invoice_filler.report import generate_mapping_doc  # noqa: E402
from invoice_filler.template import load_templates, TemplateConfig  # noqa: E402

SHOW_FIELDS = ["invoice_type", "invoice_number", "issue_date", "buyer_name", "buyer_tax_id",
               "seller_name", "seller_tax_id", "total_amount", "total_tax",
               "total_price_tax", "total_price_tax_cn", "tax_rate", "issuer",
               "buyer_bank", "buyer_account", "seller_bank", "seller_account"]


def default_template() -> TemplateConfig:
    tpls = load_templates()
    for t in tpls:
        if t.id == "template1":
            return t
    if tpls:
        return tpls[0]
    raise SystemExit("未找到任何模板配置（templates/ 目录）")


def main() -> None:
    ap = argparse.ArgumentParser(description="发票批量识别填表（命令行）")
    ap.add_argument("files", nargs="+", help="发票 PDF/图片文件")
    ap.add_argument("--template", help="模板 JSON 路径，默认模板1")
    ap.add_argument("--out", help="输出 xlsx 路径")
    ap.add_argument("--dump", action="store_true", help="打印完整字段")
    args = ap.parse_args()

    tpl = TemplateConfig.load(args.template) if args.template else default_template()
    print(f"使用模板：{tpl.name}（{tpl.id}），共 {len(tpl.columns)} 列")

    records = []
    for f in args.files:
        print(f"\n=== {os.path.basename(f)} ===")
        rec = parse_invoice(f)
        records.append(rec)
        info = {k: rec.get_field(k) for k in SHOW_FIELDS if rec.get_field(k)}
        info["status"] = rec.status
        if rec.warnings:
            info["warnings"] = rec.warnings
        if args.dump:
            info["items"] = [it.__dict__ for it in rec.items]
        print(json.dumps(info, ensure_ascii=False, indent=1))

    ok = sum(1 for r in records if r.status == "ok")
    print(f"\n识别完成：{ok}/{len(records)} 成功")

    if args.out:
        stats = export(records, tpl, args.out)
        md = generate_mapping_doc(tpl, records)
        md_path = os.path.splitext(args.out)[0] + "_字段映射说明.md"
        with open(md_path, "w", encoding="utf-8") as fh:
            fh.write(md)
        print(f"已导出：{stats['out_path']}")
        print(f"字段映射说明：{md_path}")
        print(f"统计：{stats}")


if __name__ == "__main__":
    main()
