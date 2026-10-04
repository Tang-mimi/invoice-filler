# -*- coding: utf-8 -*-
"""开发冒烟测试（自足数据，不依赖用户文件）：
1. 模板自动匹配 vs 模板1
2. 解析器直测（单明细/多明细/负数金额/换行品名）
3. GUI 流程：拖拽导入（模拟 <<Drop>>）→ 识别 → 预览 → 人工修正 → 导出
4. OCR 兜底路径（PDF→图片→RapidOCR）

注：整个测试只创建一个 Tk 实例，并在结尾用 os._exit 跳过解释器对 Tk/tkdnd
残留对象的清理，避免 Tcl_AsyncDelete 收尾竞争导致偶发假失败。
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---------------- 合成数电票 PDF ----------------
ITEMS1 = [
    {"name": "*非金属矿物制品*纤维增", "name2": "强树脂切割片", "spec": "125*1.0*22",
     "unit": "片", "qty": "500", "price": "0.9699115044248",
     "amount": "484.96", "rate": "13%", "tax": "63.04"},
]
ITEMS2 = [
    {"name": "*非金属矿物制品*纤维增", "name2": "强树脂切割片", "spec": "125*1.0*22",
     "unit": "片", "qty": "400", "price": "0.929203539823",
     "amount": "371.68", "rate": "13%", "tax": "48.32"},
    {"name": "*非金属矿物制品*纤维增", "name2": "强树脂切割片", "spec": "100*2.0*16",
     "unit": "片", "qty": "800", "price": "0.7256637168142",
     "amount": "580.53", "rate": "13%", "tax": "75.47"},
    {"name": "*非金属矿物制品*纤维增", "name2": "强树脂切割片", "spec": "405*3.2*32A",
     "unit": "片", "qty": "400", "price": "8.4070796460177",
     "amount": "3362.83", "rate": "13%", "tax": "437.17"},
    {"name": "*非金属矿物制品*纤维增", "name2": "强树脂切割片", "spec": "",
     "unit": "", "qty": "", "price": "", "amount": "-210.18", "rate": "13%", "tax": "-27.32"},
]
TOTALS = {
    "26322000008120871196": ("484.96", "63.04", "548.00", "伍佰肆拾捌圆整"),
    "26322000008121314326": ("4104.86", "533.64", "4638.50", "肆仟陆佰叁拾捌圆伍角整"),
}
BUYER1 = "测试购方贸易有限公司"


def make_invoice_pdf(path: str, number: str, items: list) -> None:
    """按数电票版式绘制一张带文字层的测试发票。"""
    import fitz

    amount, tax, total, cn = TOTALS[number]
    doc = fitz.open()
    page = doc.new_page(width=760, height=520)

    def T(x, y, s, size=9, ascii_only=False):
        # 真实票面的数字/字母列是半宽字体；中文列用 CJK 字体
        page.insert_text((x, y), s, fontsize=size,
                         fontname="helv" if ascii_only else "china-ss")

    T(240, 40, "电子发票（增值税专用发票）", 14)
    T(520, 30, f"发票号码：{number}")
    T(520, 52, "开票日期：2026年09月30日")
    T(50, 95, f"名称：{BUYER1}")
    T(50, 115, "统一社会信用代码/纳税人识别号：91460000MA5TKFKY8F")
    T(420, 95, "名称：测试销方制造有限公司")
    T(420, 115, "统一社会信用代码/纳税人识别号：9132040268963116XH")
    T(50, 150, "项目名称"); T(210, 150, "规格型号"); T(300, 150, "单 位")
    T(345, 150, "数 量"); T(400, 150, "单 价"); T(480, 150, "金 额")
    T(540, 150, "税率/征收率"); T(620, 150, "税 额")
    y = 178
    for it in items:
        T(50, y, it["name"])
        T(50, y + 13, it["name2"])
        if it["spec"]:
            T(210, y, it["spec"], ascii_only=True)
        if it["unit"]:
            T(300, y, it["unit"])
        if it["qty"]:
            T(345, y, it["qty"], ascii_only=True)
        if it["price"]:
            T(400, y, it["price"], ascii_only=True)
        T(480, y, it["amount"], ascii_only=True)
        T(540, y, it["rate"], ascii_only=True)
        T(620, y, it["tax"], ascii_only=True)
        y += 28
    y += 14
    T(50, y, "合        计")
    T(480, y, f"￥{amount}")
    T(620, y, f"￥{tax}")
    y += 30
    T(50, y, "价税合计（大写）")
    T(180, y, cn)
    T(450, y, f"（小写）￥{total}")
    y += 30
    T(50, y, "购方开户银行:测试银行文昌支行;  银行账号:1014392100000187;")
    T(50, y + 14, "销方开户银行:测试银行常州支行;  银行账号:3204215501201000516181")
    T(50, y + 50, "开票人：孙凌云")
    doc.save(path)
    doc.close()


# ---------------- 各项测试 ----------------
def _template_xlsx() -> str:
    """优先用本地真实模板，开源发布环境回退到示例模板。"""
    for rel in ("填表模板.xlsx", os.path.join("示例模板", "填表模板.xlsx")):
        p = os.path.join(BASE, rel)
        if os.path.exists(p):
            return p
    raise FileNotFoundError("未找到 填表模板.xlsx 或 示例模板/填表模板.xlsx")


def test_template_matching():
    from invoice_filler.template import build_template_from_workbook
    tpl = build_template_from_workbook(_template_xlsx(), name="匹配测试")
    print("[模板自动匹配]", flush=True)
    for c in tpl.columns:
        print(f"  {c.col}(跨{c.colspan}) {c.header} -> {c.field}", flush=True)
    got = {c.header: c.field for c in tpl.columns}
    expect = {"序号": "_row_number", "开票日期": "issue_date", "发票号码": "invoice_number",
              "客户名称": "buyer_name", "交易模式": None, "金额": "total_amount",
              "税额": "total_tax", "价税合计": "total_price_tax"}
    assert got == expect, f"自动匹配与模板1不一致: {got}"
    assert tpl.columns[3].colspan == 3, "客户名称应检测到 E:G 合并"
    print("  OK 与模板1一致，合并单元格 E:G 检测正确\n", flush=True)


def test_parser_direct(pdfs):
    from invoice_filler.parser import parse_invoice
    for pdf, items in zip(pdfs, (ITEMS1, ITEMS2)):
        rec = parse_invoice(pdf)
        number = rec.get_field("invoice_number")
        amount, tax, total, cn = TOTALS[number]
        print(f"[解析] {number}: {rec.summary()} status={rec.status} items={len(rec.items)}", flush=True)
        assert rec.status == "ok", rec.warnings
        assert rec.get_field("buyer_name") == BUYER1
        assert rec.get_field("seller_name") == "测试销方制造有限公司"
        assert rec.get_field("buyer_tax_id") == "91460000MA5TKFKY8F"
        assert rec.get_field("issue_date") == "2026-09-30"
        assert rec.get_field("total_amount") == amount
        assert rec.get_field("total_tax") == tax
        assert rec.get_field("total_price_tax") == total
        assert rec.get_field("total_price_tax_cn") == cn
        assert rec.get_field("tax_rate") == "13%"
        assert rec.get_field("buyer_account") == "1014392100000187"
        assert rec.get_field("seller_account") == "3204215501201000516181"
        assert len(rec.items) == len(items), rec.items
        assert rec.items[0].name == "*非金属矿物制品*纤维增强树脂切割片"


def test_app_flow(pdf1, pdf2):
    """拖拽导入 + GUI 识别/修正/导出（共用一个 Tk 实例）。"""
    from ui import app as appmod
    from ui.app import InvoiceFillerApp
    assert appmod._DND, "tkinterdnd2 未生效，拖拽不可用"
    app = InvoiceFillerApp()
    app.update()

    # 验证真实事件链：<<Drop>> 绑定脚本确实通过 dnd_bind 注册到了根窗口
    script = app.tk.call("bind", str(app), "<<Drop>>")
    assert script and "%D" in str(script), f"<<Drop>> 绑定异常: {script}"

    class Ev:
        pass

    def drop(data):
        ev = Ev()
        ev.data = data
        app._on_drop(ev)

    tmpd = tempfile.mkdtemp()
    try:
        p_space = os.path.join(tmpd, "带 空格 的 文件.pdf")
        shutil.copy(pdf1, p_space)
        shutil.copy(pdf2, os.path.join(tmpd, "c.pdf"))

        def dpath(p):  # tkdnd 在 Windows 上的真实格式：正斜杠、含空格的项用花括号包裹
            return "{" + p.replace("\\", "/") + "}"

        drop(f'{dpath(pdf1)} {dpath(p_space)}')     # 直接拖两个文件（一个带空格路径）
        assert len(app.files) == 2, app.files
        with open(os.path.join(tmpd, "说明.md"), "w", encoding="utf-8") as f:
            f.write("x")
        drop(tmpd.replace("\\", "/"))               # 拖文件夹（1 个支持 + 1 个不支持）
        assert len(app.files) == 3, app.files
        drop(tmpd.replace("\\", "/"))               # 重复拖入应去重
        assert len(app.files) == 3
        drop(dpath(os.path.join(BASE, "README.md")))  # 不支持的类型
        assert len(app.files) == 3
        print(f"[拖拽导入] OK，共 {len(app.files)} 个文件", flush=True)

        # 重置为标准两票，走识别 → 预览 → 修正 → 导出
        app.files = [pdf1, pdf2]
        app.records = {}
        app.refresh_file_tree()
        app.rebuild_preview_rows()
        app.start_recognition()
        for _ in range(2000):
            app.update()
            if not app.running:
                break
            app.after(50)
            app.update()
        assert not app.running, "识别未结束"
        ok = sum(1 for r in app.records.values() if r.status == "ok")
        print(f"[GUI识别] {ok}/2 成功", flush=True)
        assert ok == 2
        rows = [app.preview.item(i)["values"] for i in app.preview.get_children()]
        assert rows[0][4] == BUYER1, rows[0]
        assert rows[0][5] == "", "交易模式应留白"
        rec1 = app.records[0]
        assert rec1.get_field("total_amount") == TOTALS[rec1.get_field("invoice_number")][0]
        expect_items = {"26322000008120871196": 1, "26322000008121314326": 4}
        assert len(rec1.items) == expect_items[rec1.get_field("invoice_number")]

        # 人工修正
        iid = app.preview.get_children()[0]
        app.commit_edit(iid, "buyer_name", "手动改的名字有限公司")
        assert app.records[0].get_field("buyer_name") == "手动改的名字有限公司"
        app.records[0].set_field("buyer_name", "")
        assert app.records[0].get_field("buyer_name") == BUYER1

        # 导出
        out = os.path.join(BASE, "输出", "_smoke_gui.xlsx") if os.path.isdir(
            os.path.join(BASE, "输出")) else os.path.join(tmpd, "_smoke_gui.xlsx")
        from invoice_filler.exporter import export
        stats = export(list(app.records.values()), app.current_template, out)
        assert stats["total"] == 2
        if os.path.exists(out):
            os.remove(out)
        print("  导出 OK", flush=True)
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)
        try:
            app.destroy()
        except Exception:
            pass


def test_ocr_fallback(pdf1):
    """把合成发票渲染成图片走 OCR 路径，验证免密钥 OCR 兜底可用。"""
    import fitz
    from invoice_filler.parser import parse_invoice
    img_path = os.path.join(tempfile.mkdtemp(), "_ocr_test.png")
    doc = fitz.open(pdf1)
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72), alpha=False)
    pix.save(img_path)
    doc.close()
    rec = parse_invoice(img_path)
    print("[OCR兜底]", json.dumps({k: rec.get_field(k) for k in
          ("invoice_number", "issue_date", "buyer_name", "total_price_tax", "status")},
          ensure_ascii=False), flush=True)
    assert rec.invoice_number == "26322000008120871196", rec.invoice_number
    assert rec.buyer_name, "OCR 未识别到购买方名称"
    assert rec.total_price_tax, "OCR 未识别到价税合计"
    shutil.rmtree(os.path.dirname(img_path), ignore_errors=True)


def main():
    tmp = tempfile.mkdtemp(prefix="invoice_test_")
    code = 0
    try:
        pdf1 = os.path.join(tmp, "invoice_single.pdf")
        pdf2 = os.path.join(tmp, "invoice_multi.pdf")
        make_invoice_pdf(pdf1, "26322000008120871196", ITEMS1)
        make_invoice_pdf(pdf2, "26322000008121314326", ITEMS2)
        test_template_matching()
        test_parser_direct([pdf1, pdf2])
        test_app_flow(pdf1, pdf2)
        test_ocr_fallback(pdf1)
        print("\n全部冒烟测试通过 ✓", flush=True)
    except Exception:
        traceback.print_exc()
        code = 1
    shutil.rmtree(tmp, ignore_errors=True)
    sys.stdout.flush()
    sys.stderr.flush()
    # 跳过解释器对 Tk/tkdnd 残留对象的清理（多根窗口收尾竞争会偶发崩溃）
    os._exit(code)


if __name__ == "__main__":
    main()
