# -*- coding: utf-8 -*-
"""发票文件读取：PDF 优先抽取文字层（精确），扫描件/图片走本地 OCR（RapidOCR，免密钥）。

统一输出为词块列表 Word(x0,y0,x1,y1,text)，供解析器按行/区域使用，
避免 cv2.imread 等在中文路径下的问题（自行解码字节）。
"""
from __future__ import annotations

import os
from typing import List, NamedTuple, Optional

import numpy as np

TEXT_LAYER_MIN_CHARS = 20      # 文字层少于该字符数视为扫描件，走 OCR
OCR_RENDER_DPI = 300           # 扫描件渲染分辨率


class Word(NamedTuple):
    x0: float
    y0: float
    x1: float
    y1: float
    text: str


class PageData(NamedTuple):
    words: List[Word]
    width: float
    height: float
    engine: str                 # "text-layer" / "ocr"


_ocr_engine = None


def get_ocr_engine():
    """惰性加载 RapidOCR，首次加载约需数秒。"""
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    return _ocr_engine


def ocr_image_array(img_bgr: np.ndarray) -> List[Word]:
    """对 BGR 图像做 OCR，返回词块。兼容 rapidocr 1.x 的不同返回结构。"""
    engine = get_ocr_engine()
    result, _elapse = engine(img_bgr)
    words: List[Word] = []
    if not result:
        return words
    for item in result:
        try:
            box, text, score = item[0], item[1], item[2]
        except Exception:
            continue
        if not text or float(score) < 0.4:
            continue
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        words.append(Word(float(min(xs)), float(min(ys)), float(max(xs)), float(max(ys)), str(text)))
    return words


def read_image_words(path: str) -> List[Word]:
    """读取图片文件（png/jpg/bmp/tiff 等），中文/特殊路径安全。"""
    import cv2
    with open(path, "rb") as f:
        data = np.frombuffer(f.read(), dtype=np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError("无法解码图片文件")
    return ocr_image_array(img)


def render_pdf_page(page) -> np.ndarray:
    """把 PDF 页渲染成 300dpi 的 BGR ndarray。"""
    import cv2
    zoom = OCR_RENDER_DPI / 72.0
    pix = page.get_pixmap(matrix=__import__("fitz").Matrix(zoom, zoom), alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
    elif pix.n == 3:
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    else:
        img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
    return img


def pick_best_page(doc) -> int:
    """发票 PDF 一般单页；多页时取文字最多的一页。"""
    best, best_len = 0, -1
    for i in range(doc.page_count):
        try:
            n = len(doc[i].get_text("text"))
        except Exception:
            n = 0
        if n > best_len:
            best, best_len = i, n
    return best


def load_page(path: str) -> PageData:
    """读取 PDF 或图片，返回词块页面数据。PDF 无文字层时自动渲染后 OCR。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        import fitz
        doc = fitz.open(path)
        try:
            if doc.needs_pass and not doc.authenticate(""):
                raise ValueError("PDF 已加密，无法读取")
            idx = pick_best_page(doc)
            page = doc[idx]
            text = page.get_text("text") or ""
            if len(text.strip()) >= TEXT_LAYER_MIN_CHARS:
                words = [
                    Word(x0, y0, x1, y1, w)
                    for x0, y0, x1, y1, w, *_ in page.get_text("words")
                    if w.strip()
                ]
                return PageData(words, page.rect.width, page.rect.height, "text-layer")
            img = render_pdf_page(page)
            return PageData(ocr_image_array(img), img.shape[1], img.shape[0], "ocr")
        finally:
            doc.close()
    else:
        words = read_image_words(path)
        if not words:
            return PageData([], 0, 0, "ocr")
        w = max(wd.x1 for wd in words)
        h = max(wd.y1 for wd in words)
        return PageData(words, w, h, "ocr")


def page_full_text(page: PageData) -> str:
    """按行序拼接的整页文本（调试与兜底用）。"""
    return "\n".join(band_text(b) for b in group_bands(page.words)[0])


# ---------------------------------------------------------------------------
# 行聚合：把词块按 y 坐标聚成"视觉行"（band），文字层与 OCR 通用
# ---------------------------------------------------------------------------
class Band(NamedTuple):
    y: float            # 行中心 y
    words: List[Word]

    @property
    def text(self) -> str:
        return " ".join(w.text for w in sorted(self.words, key=lambda w: w.x0))


def group_bands(words: List[Word], y_tol_ratio: float = 0.6):
    """按 y 中心聚类成行。返回 (bands, median_height)。"""
    if not words:
        return [], 10.0
    hs = sorted((w.y1 - w.y0) for w in words)
    med_h = hs[len(hs) // 2] or 10.0
    tol = med_h * y_tol_ratio
    ordered = sorted(words, key=lambda w: ((w.y0 + w.y1) / 2, w.x0))
    bands: List[List[Word]] = []
    centers: List[float] = []
    for w in ordered:
        c = (w.y0 + w.y1) / 2
        if not bands:
            bands.append([w])
            centers.append(c)
            continue
        if abs(c - centers[-1]) <= tol:
            bands[-1].append(w)
            centers[-1] = sum((x.y0 + x.y1) / 2 for x in bands[-1]) / len(bands[-1])
        else:
            bands.append([w])
            centers.append(c)
    result = [Band(sum((x.y0 + x.y1) / 2 for x in ws) / len(ws), ws) for ws in bands]
    return result, med_h


def words_in_y_range(words: List[Word], y0: float, y1: float) -> List[Word]:
    return [w for w in words if y0 <= (w.y0 + w.y1) / 2 <= y1]


def band_text(band: Band) -> str:
    return band.text
