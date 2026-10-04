# -*- coding: utf-8 -*-
"""数电票（电子发票）解析器：把词块还原为发票规范字段。

对文字层与 OCR 词块使用同一套"行聚合 + 锚点定位"逻辑：
1. 按标签（发票号码/开票日期/名称/纳税人识别号…）在视觉行内定位取值；
2. 明细表用表头列锚点按 x 坐标切列，按 y 聚行，处理名称换行；
3. 每个字段失败时回退到整页正则，仍未命中则留空并记录警告。
"""
from __future__ import annotations

import os
import re
from typing import Dict, List, Optional, Tuple

from .fields import SUPPORTED_TITLE_KEYWORDS
from .models import InvoiceData, Item, ST_FAILED, ST_UNSUPPORTED
from .reader import Band, PageData, group_bands, load_page

RE_NUM_LABEL = re.compile(r"发\s*票\s*号\s*码\s*[:：]")
RE_NUMBER_VAL = re.compile(r"\d{10,22}")
RE_DATE_LABEL = re.compile(r"开\s*票\s*日\s*期\s*[:：]")
RE_DATE_VAL = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
RE_NAME_LABEL = re.compile(r"名\s*称\s*[:：]")
RE_TAXID = re.compile(
    r"(?:统一\s*社\s*会\s*信\s*用\s*代\s*码\s*[/／]?\s*纳\s*税\s*人\s*识\s*别\s*号"
    r"|纳\s*税\s*人\s*识\s*别\s*号|统一\s*社\s*会\s*信\s*用\s*代\s*码)\s*[:：]\s*([0-9A-Z]{15,22})"
)
RE_PRICE_TAX_LC = re.compile(r"[（(]\s*小\s*写\s*[)）]\s*[¥￥]?\s*(-?[\d,]+(?:\.\d+)?)")
RE_PRICE_TAX_CAP = re.compile(r"[（(]\s*大\s*写\s*[)）](.+?)[（(]\s*小\s*写\s*[)）]", re.S)
RE_NUMERIC_TOKEN = re.compile(r"[¥￥]?\s*(-?[\d,]+(?:\.\d+)?)")
RE_CJK_SPACE = re.compile(r"(?<=[\u4e00-\u9fff（）《》*％%])\s+(?=[\u4e00-\u9fff（）《》*％%])")

ITEM_HEADERS = [
    ("name", re.compile(r"项目\s*名\s*称|货\s*物\s*或\s*应\s*税\s*劳\s*务|货\s*物\s*或\s*服\s*务\s*名\s*称")),
    ("spec", re.compile(r"规\s*格\s*型\s*号")),
    ("unit", re.compile(r"单\s*位")),
    ("qty", re.compile(r"数\s*量")),
    ("price", re.compile(r"单\s*价")),
    ("amount", re.compile(r"金\s*额")),
    ("rate", re.compile(r"税\s*率\s*[/／]?\s*征\s*收\s*率|税\s*率")),
    ("tax", re.compile(r"税\s*额")),
]

NUMERIC_COLS = ("qty", "price", "amount", "tax")

# 竖排"购买方信息/销售方信息"标签可能混进行首/行尾，剔除这些孤字
LABEL_STRAY_CHARS = "购买方信息销售"


def strip_label_stray(text: str) -> str:
    while text and text[-1] in LABEL_STRAY_CHARS and len(text) > 4:
        text = text[:-1]
    return text


CJK_RANGES = ((0x2E80, 0x9FFF), (0xF900, 0xFAFF), (0xFF00, 0xFFEF))


def is_cjk_char(ch: str) -> bool:
    o = ord(ch)
    return any(a <= o <= b for a, b in CJK_RANGES)


def split_char_runs(text: str) -> List[str]:
    """把字符串切成 CJK / 非CJK 的连续段，用于拆开跨列粘连的词。"""
    runs: List[str] = []
    for ch in text:
        c = is_cjk_char(ch)
        if runs and is_cjk_char(runs[-1][-1]) == c:
            runs[-1] += ch
        else:
            runs.append(ch)
    return runs


def find_totals_band(bands):
    """合计行：去掉空白后以"合计"开头（排除"价税合计…"）。"""
    for b in bands:
        t = re.sub(r"\s+", "", b.text)
        if t.startswith("合计") and "价税合计" not in t:
            return b
    return None


def clean_number(text: str) -> str:
    t = (text or "").replace("¥", "").replace("￥", "").replace(",", "").replace(" ", "")
    t = t.strip("()（）")
    return t


def is_numeric_token(text: str) -> Optional[str]:
    t = clean_number(text)
    if t and re.fullmatch(r"-?\d+(\.\d+)?", t):
        return t
    return None


def cjk_clean(text: str) -> str:
    """去掉中文字符之间因分行/OCR 产生的空格。"""
    return RE_CJK_SPACE.sub("", text or "").strip()


def dedup_space(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


# ---------------------------------------------------------------------------
def band_parts(band: Band) -> Tuple[str, List[Tuple[int, int, "Word"]]]:
    """行文本 + 每个词的字符区间，用于把匹配位置映射回 x 坐标。"""
    from .reader import Word

    words = sorted(band.words, key=lambda w: w.x0)
    parts: List[Tuple[int, int, Word]] = []
    pos = 0
    for w in words:
        if parts:
            pos += 1  # 词间空格
        parts.append((pos, pos + len(w.text), w))
        pos += len(w.text)
    text = " ".join(w.text for w in words)
    return text, parts


def word_x_at(parts, char_pos: int) -> float:
    for s, e, w in parts:
        if s <= char_pos < e:
            return w.x0
    return parts[-1][2].x0 if parts else 0.0


def label_values_in_band(band: Band, label_re: re.Pattern) -> List[Tuple[float, float, str]]:
    """在行内按标签切值：返回 [(标签x, 行y, 值文本), ...]。"""
    text, parts = band_parts(band)
    matches = list(label_re.finditer(text))
    out = []
    for i, m in enumerate(matches):
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        val = dedup_space(text[start:end])
        if val:
            out.append((word_x_at(parts, m.start()), band.y, val))
    return out


class InvoiceParser:
    def __init__(self, page: PageData):
        self.page = page
        self.bands, self.med_h = group_bands(page.words)
        self.texts = [b.text for b in self.bands]
        self.full_text = "\n".join(self.texts)
        self.mid_x = page.width / 2 if page.width else 1e9

    # ---------- 基础取值 ----------
    def _label_same_band(self, label_re, val_re) -> Optional[str]:
        for band in self.bands:
            text = band.text
            m = label_re.search(text)
            if not m:
                continue
            rest = text[m.end():]
            vm = val_re.search(rest)
            if vm:
                return vm.group(1) if vm.groups() else vm.group(0)
        return None

    def _full_fallback(self, regex: re.Pattern, group: int = 1) -> Optional[str]:
        m = regex.search(self.full_text)
        return m.group(group) if m else None

    # ---------- 各字段 ----------
    def extract_title(self) -> str:
        for band in self.bands[:8]:
            t = re.sub(r"\s+", "", band.text)
            if "发票" not in t:
                continue
            cut = re.split(r"发票号码|开票日期", t)[0]
            if "发票" in cut and len(cut) <= 40:
                return cut
        t0 = re.sub(r"\s+", "", (self.texts[0] if self.texts else ""))
        return re.split(r"发票号码|开票日期", t0)[0]

    def extract_number(self) -> Optional[str]:
        v = self._label_same_band(RE_NUM_LABEL, RE_NUMBER_VAL)
        if v:
            return v
        m = self._full_fallback(re.compile(r"发\s*票\s*号\s*码\s*[:：]\s*(\d{10,22})"))
        if m:
            return m
        # 兜底：数电票号码固定20位
        m20 = re.search(r"(?<!\d)(\d{20})(?!\d)", self.full_text)
        return m20.group(1) if m20 else None

    def extract_date(self) -> Optional[str]:
        m = re.search(r"开\s*票\s*日\s*期\s*[:：]?\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日",
                      self.full_text)
        if not m:
            for band in self.bands:
                if RE_DATE_LABEL.search(band.text):
                    m2 = RE_DATE_VAL.search(band.text)
                    if m2:
                        m = m2
                        break
        if not m:
            return None
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            import datetime
            datetime.date(y, mo, d)
        except ValueError:
            return None
        return f"{y:04d}-{mo:02d}-{d:02d}"

    def extract_names(self) -> Tuple[Optional[str], Optional[str]]:
        cands: List[Tuple[float, float, str]] = []
        for band in self.bands:
            cands.extend(label_values_in_band(band, RE_NAME_LABEL))
        cands.sort(key=lambda c: (c[1], c[0]))
        if not cands:
            return None, None
        left = [c for c in cands if c[0] < self.mid_x]
        right = [c for c in cands if c[0] >= self.mid_x]
        buyer = cjk_clean(strip_label_stray(left[0][2])) if left else None
        seller = cjk_clean(strip_label_stray(right[0][2])) if right else None
        if (buyer is None or seller is None) and len(cands) >= 2:
            buyer = buyer or cjk_clean(strip_label_stray(cands[0][2]))
            seller = seller or cjk_clean(strip_label_stray(cands[1][2]))
        return buyer, seller

    def extract_taxids(self) -> Tuple[Optional[str], Optional[str]]:
        cands: List[Tuple[float, float, str]] = []
        for band in self.bands:
            text, parts = band_parts(band)
            for m in RE_TAXID.finditer(text):
                cands.append((word_x_at(parts, m.start()), band.y, m.group(1)))
        cands.sort(key=lambda c: (c[1], c[0]))
        if not cands:
            return None, None
        left = [c for c in cands if c[0] < self.mid_x]
        right = [c for c in cands if c[0] >= self.mid_x]
        buyer = left[0][2] if left else None
        seller = right[0][2] if right else None
        if (buyer is None or seller is None) and len(cands) >= 2:
            buyer = buyer or cands[0][2]
            seller = seller or cands[1][2]
        return buyer, seller

    def find_band(self, pattern: str, start: int = 0) -> Optional[Band]:
        rx = re.compile(pattern)
        for band in self.bands[start:]:
            if rx.search(band.text):
                return band
        return None

    def band_index(self, band: Optional[Band]) -> int:
        if band is None:
            return -1
        for i, b in enumerate(self.bands):
            if b is band:
                return i
        return -1

    def numeric_tokens(self, band: Band) -> List[Tuple[float, str]]:
        toks = []
        for w in sorted(band.words, key=lambda w: w.x0):
            n = is_numeric_token(w.text)
            if n is not None:
                toks.append((w.x0, n))
        return toks

    def extract_totals(self) -> Tuple[Optional[str], Optional[str]]:
        band = find_totals_band(self.bands)
        if band is not None:
            toks = self.numeric_tokens(band)
            if len(toks) < 2:
                # 金额与税额可能被分到相邻行，向上下各找一行补齐
                idx = self.band_index(band)
                for j in (idx - 1, idx + 1):
                    if 0 <= j < len(self.bands):
                        toks.extend(self.numeric_tokens(self.bands[j]))
                toks.sort(key=lambda t: t[0])
            if len(toks) >= 2:
                return toks[0][1], toks[1][1]
            if len(toks) == 1:
                return toks[0][1], None
        return None, None

    def extract_price_tax(self) -> Tuple[Optional[str], Optional[str]]:
        """返回 (小写金额, 大写文本)。"""
        lc = cap = None
        band = self.find_band(r"小\s*写")
        if band is not None:
            toks = self.numeric_tokens(band)
            if toks:
                lc = toks[-1][1]
            text = band.text
            m1 = re.search(r"[（(]\s*大\s*写\s*[)）]", text)
            m2 = re.search(r"[（(]\s*小\s*写\s*[)）]", text)
            if m1:
                seg = text[m1.end(): m2.start() if m2 else len(text)]
                seg = re.sub(r"[¥￥⊗※○◎〇xX×]", "", seg)
                seg = re.sub(r"[\d,\.\s（）()]", "", seg)
                cap = seg.strip() or None
        if lc is None:
            m = self._full_fallback(RE_PRICE_TAX_LC)
            lc = clean_number(m) if m else None
        if cap is None:
            m = self._full_fallback(RE_PRICE_TAX_CAP)
            if m:
                seg = re.sub(r"[¥￥⊗※○◎〇xX×]", "", m.group(1))
                seg = re.sub(r"[\d,\.\s（）()]", "", seg)
                cap = seg.strip() or None
        return lc, cap

    def extract_issuer(self) -> Optional[str]:
        for band in self.bands:
            m = re.search(r"开\s*票\s*人\s*[:：]", band.text)
            if m:
                val = dedup_space(band.text[m.end():])
                val = re.sub(r"[（）()]", "", val).strip()
                if val:
                    return cjk_clean(val)
        m = self._full_fallback(re.compile(r"开\s*票\s*人\s*[:：]\s*(\S+)"))
        return m if m else None

    def extract_remark(self, y_start: float, y_end: float) -> str:
        lines = []
        for band in self.bands:
            if y_start < band.y < y_end:
                t = band.text.strip()
                if re.fullmatch(r"[备]\s*[注]?", t) or t in ("备", "注"):
                    continue
                if t:
                    lines.append(t)
        return cjk_clean(dedup_space(" ".join(lines)))

    # ---------- 明细表 ----------
    def header_anchors(self, header_band: Band) -> List[Tuple[str, float]]:
        text, parts = band_parts(header_band)
        found: List[Tuple[str, float]] = []
        used: List[Tuple[int, int]] = []
        for key, rx in ITEM_HEADERS:
            for m in rx.finditer(text):
                span = (m.start(), m.end())
                if any(not (m.end() <= s or m.start() >= e) for s, e in used):
                    continue
                used.append(span)
                found.append((key, word_x_at(parts, m.start())))
                break
        found.sort(key=lambda a: a[1])
        return found

    def extract_items(self, y_top: float, y_bottom: float) -> List[Item]:
        header_band = self.find_band(r"项\s*目\s*名\s*称|货\s*物\s*或\s*服\s*务\s*名\s*称")
        if header_band is None:
            return []
        anchors = self.header_anchors(header_band)
        if len(anchors) < 3:
            return []
        region = [w for b in self.bands if y_top < b.y < y_bottom for w in b.words]
        if not region:
            return []
        rows, _ = group_bands(region, 0.5)

        # 列边界：相邻锚点中点
        bounds = []
        for i, (key, x) in enumerate(anchors):
            lo = (anchors[i - 1][1] + x) / 2 if i else x - 60
            hi = (x + anchors[i + 1][1]) / 2 if i + 1 < len(anchors) else x + 400
            bounds.append((key, lo, hi))

        def cols_for_word(w) -> List[Tuple[str, str]]:
            """词 → [(列, 文本)]。
            整词（含2pt容差）落在单列边界内 → 整词归该列；
            跨列粘连词（OCR 常见）按 CJK/ASCII 连续段切分，段的起点越过
            下一列锚点（5pt 容差）才切换列，避免长品名溢出列宽被误切。"""
            for key, lo, hi in bounds:
                if lo - 2 <= w.x0 and w.x1 <= hi + 2:
                    return [(key, w.text)]
            n = len(w.text)
            if n == 0:
                return []
            anchor_xs = [x for _, x in anchors]
            step = (w.x1 - w.x0) / n
            out: List[Tuple[str, str]] = []
            pos = 0
            for run in split_char_runs(w.text):
                rx0 = w.x0 + step * pos
                pos += len(run)
                j = 0
                for idx, ax in enumerate(anchor_xs):
                    if rx0 >= ax - 5:
                        j = idx
                out.append((anchors[j][0], run))
            return out

        raw_rows: List[Dict[str, str]] = []
        for row in rows:
            cols: Dict[str, List[str]] = {}
            for w in sorted(row.words, key=lambda w: w.x0):
                for key, seg in cols_for_word(w):
                    cols.setdefault(key, []).append(seg)
            raw_rows.append({k: "".join(v) for k, v in cols.items() if v})

        # 合并"仅名称/规格"的换行行到相邻数据行
        merged: List[Dict[str, str]] = []
        for r in raw_rows:
            has_num = any(is_numeric_token(r.get(c, "")) for c in NUMERIC_COLS)
            if has_num or not merged:
                merged.append(dict(r))
                continue
            name_part = r.get("name", "")
            spec_part = r.get("spec", "")
            if merged:
                prev = merged[-1]
                prev_is_data = any(is_numeric_token(prev.get(c, "")) for c in NUMERIC_COLS)
                nxt = None
                idx = raw_rows.index(r)
                for later in raw_rows[idx + 1:]:
                    if any(is_numeric_token(later.get(c, "")) for c in NUMERIC_COLS):
                        nxt = later
                        break
                if not prev_is_data and nxt is not None:
                    nxt["name"] = name_part + (nxt.get("name", "") or "")
                    if spec_part:
                        nxt["spec"] = spec_part + (nxt.get("spec", "") or "")
                elif prev_is_data:
                    prev["name"] = (prev.get("name", "") or "") + name_part
                    if spec_part:
                        prev["spec"] = (prev.get("spec", "") or "") + spec_part
                else:
                    merged.append(dict(r))

        items: List[Item] = []
        for r in merged:
            name = cjk_clean(r.get("name", ""))
            if name in ("合", "计", "合计", "小计"):   # 合计行残余
                continue
            if not name and not any(is_numeric_token(r.get(c, "")) for c in NUMERIC_COLS):
                continue
            items.append(Item(
                name=name,
                spec=cjk_clean(r.get("spec", "")),
                unit=cjk_clean(r.get("unit", "")),
                qty=clean_number(r.get("qty", "")),
                price=clean_number(r.get("price", "")),
                amount=clean_number(r.get("amount", "")),
                tax_rate=r.get("rate", ""),
                tax_amount=clean_number(r.get("tax", "")),
            ))
        return items

    # ---------- 主流程 ----------
    def parse(self, source_file: str) -> InvoiceData:
        rec = InvoiceData(source_file=source_file, engine=self.page.engine)
        if not self.bands:
            rec.status = ST_FAILED
            rec.warnings.append("未检测到任何文字内容")
            return rec

        title = self.extract_title()
        rec.invoice_type = title
        supported = any(k in title for k in SUPPORTED_TITLE_KEYWORDS)
        if not supported:
            if re.search(r"(?<!\d)\d{20}(?!\d)", self.full_text) and "识别号" in self.full_text:
                rec.warnings.append("发票标题未识别到，按数电票结构解析")
            else:
                rec.status = ST_UNSUPPORTED
                rec.warnings.append(f"版式不支持：{title or '未识别到发票标题'}，字段留空")
                return rec

        rec.invoice_number = self.extract_number() or ""
        rec.issue_date = self.extract_date() or ""
        buyer, seller = self.extract_names()
        rec.buyer_name = buyer or ""
        rec.seller_name = seller or ""
        bt, stx = self.extract_taxids()
        rec.buyer_tax_id = bt or ""
        rec.seller_tax_id = stx or ""

        price_tax_band = self.find_band(r"小\s*写")
        totals_band = find_totals_band(self.bands)
        header_band = self.find_band(r"项\s*目\s*名\s*称|货\s*物\s*或\s*服\s*务\s*名\s*称")
        y_top = header_band.y if header_band else 0
        y_bottom = totals_band.y if totals_band is not None else (
            price_tax_band.y if price_tax_band else self.page.height)
        rec.items = self.extract_items(y_top + self.med_h * 0.4, y_bottom - self.med_h * 0.4)

        rec.total_amount, rec.total_tax = self.extract_totals()
        lc, cap = self.extract_price_tax()
        rec.total_price_tax = lc or ""
        rec.total_price_tax_cn = cap or ""

        totals_idx = self.band_index(find_totals_band(self.bands))
        pt_idx = self.band_index(price_tax_band)
        issuer_band = self.find_band(r"开\s*票\s*人")
        remark_top = (self.bands[pt_idx].y if pt_idx >= 0 else (self.bands[totals_idx].y if totals_idx >= 0 else 0))
        remark_bottom = (self.bands[self.band_index(issuer_band)].y - self.med_h * 0.4
                         if issuer_band is not None else self.page.height)
        rec.remark = self.extract_remark(remark_top, remark_bottom)

        m = re.search(r"购\s*方\s*开\s*户\s*银\s*行\s*[:：]\s*([^;；]+)", rec.remark)
        rec.buyer_bank = cjk_clean(m.group(1)) if m else ""
        m = re.search(r"销\s*方\s*开\s*户\s*银\s*行\s*[:：]\s*([^;；]+)", rec.remark)
        rec.seller_bank = cjk_clean(m.group(1)) if m else ""
        accounts = re.findall(r"银\s*行\s*账\s*号\s*[:：]\s*(\d{6,})", rec.remark)
        if accounts:
            rec.buyer_account = accounts[0]
            if len(accounts) > 1:
                rec.seller_account = accounts[1]

        rec.issuer = self.extract_issuer() or ""

        # 税率：取明细去重集合
        rates = [it.tax_rate for it in rec.items if it.tax_rate]
        seen, uniq = set(), []
        for r in rates:
            if r not in seen:
                seen.add(r)
                uniq.append(r)
        rec.tax_rate = "；".join(uniq)

        for label, val in [("发票号码", rec.invoice_number), ("开票日期", rec.issue_date),
                           ("购买方名称", rec.buyer_name), ("购买方纳税人识别号", rec.buyer_tax_id),
                           ("销售方名称", rec.seller_name), ("销售方纳税人识别号", rec.seller_tax_id),
                           ("合计金额", rec.total_amount), ("合计税额", rec.total_tax),
                           ("价税合计", rec.total_price_tax)]:
            if not val:
                rec.warnings.append(f"未识别到「{label}」")
        rec.finalize_status()
        return rec


def parse_invoice(path: str) -> InvoiceData:
    """入口：读取文件并解析。解析异常转为 failed 状态，不中断批量流程。"""
    try:
        page = load_page(path)
        return InvoiceParser(page).parse(path)
    except Exception as e:  # noqa: BLE001 - 单文件失败不能中断批量
        rec = InvoiceData(source_file=path)
        rec.status = ST_FAILED
        rec.warnings.append(f"识别失败：{type(e).__name__}: {e}")
        return rec
