"""OCR 识别与公文版面结构还原。

把图片或 PDF 里的文字行还原成"标题 / 主送机关 / 一级标题 / 正文 / 落款"这样的
公文语义块，供 Word 生成器套用格式。
"""

from __future__ import annotations

import base64
import io
import re
import statistics
from typing import Any

import cv2
import numpy as np
from PIL import Image, ImageOps
from rapidocr_onnxruntime import RapidOCR

try:  # PyMuPDF 新版包名为 pymupdf，旧版为 fitz
    import pymupdf as fitz
except ImportError:  # pragma: no cover
    import fitz

MAX_PAGES = 30
MAX_PREVIEWS = 6

# 公文里常见的序号形式
_H1 = re.compile(r"^[一二三四五六七八九十]+、")
_H2 = re.compile(r"^[（(][一二三四五六七八九十]+[)）]")
_H3 = re.compile(r"^(?:\d+[、.．](?!\d)|[（(]\d+[)）])")
_DATE = re.compile(
    r"^[〇零一二三四五六七八九O0-9]{2,4}年[〇零一二三四五六七八九十0-9]{1,2}月"
    r"[〇零一二三四五六七八九十0-9]{1,3}日$"
)

_ENGINE: RapidOCR | None = None


def _engine() -> RapidOCR:
    global _ENGINE
    if _ENGINE is None:
        _ENGINE = RapidOCR()
    return _ENGINE


def _decode(data: bytes) -> np.ndarray:
    """解码为 BGR 图，按 EXIF 摆正方向，并放大过小的图片。"""
    try:
        pil = Image.open(io.BytesIO(data))
        pil = ImageOps.exif_transpose(pil).convert("RGB")
    except Exception as exc:  # noqa: BLE001 - 用户上传内容不可控
        raise ValueError("无法读取图片，请确认文件是有效的 JPG 或 PNG") from exc

    img = np.array(pil)[:, :, ::-1]  # RGB -> BGR
    h, w = img.shape[:2]
    if max(h, w) < 1400:
        scale = 1400 / max(h, w)
        img = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC)
    return img


def analyze(image_bytes: bytes) -> dict[str, Any]:
    """识别图片并返回结构化段落。"""
    img = _decode(image_bytes)
    height, width = img.shape[:2]
    raw, _ = _engine()(img)
    rows = _rows(raw or [])
    if not rows:
        raise ValueError("未识别到文字，请确认图片清晰、方向正确")
    return {"blocks": _structure(rows, width), "width": width, "height": height}


def analyze_pdf(data: bytes) -> dict[str, Any]:
    """解析 PDF：带文字层的页直接取字，扫描页渲染成图后走 OCR。"""
    try:
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception as exc:  # noqa: BLE001 - 用户上传内容不可控
        raise ValueError("无法读取 PDF，请确认文件未损坏或未加密") from exc

    try:
        if doc.page_count == 0:
            raise ValueError("PDF 里没有页面")
        if doc.page_count > MAX_PAGES:
            raise ValueError(f"PDF 共 {doc.page_count} 页，目前最多支持 {MAX_PAGES} 页")

        blocks: list[dict[str, Any]] = []
        previews: list[str] = []
        for page in doc:
            rows, width = _page_rows(page)
            if not rows:
                continue
            if blocks:  # 页与页之间插入分页符
                blocks.append({"role": "pagebreak", "text": "", "score": 1.0})
            blocks.extend(_structure(rows, width))
            if len(previews) < MAX_PREVIEWS:
                previews.append(_preview(page))

        if not blocks:
            raise ValueError("未识别到文字，可能是空白 PDF 或扫描过于模糊")
        return {"blocks": blocks, "pages": doc.page_count, "previews": previews}
    finally:
        doc.close()


def _page_rows(page) -> tuple[list[dict[str, Any]], float]:
    """返回该页的文字行，以及行坐标所在的宽度（用于居中/缩进判断）。"""
    text = re.sub(r"[\s\u3000]", "", page.get_text("text"))
    if len(text) >= 40:  # 有文字层，直接取字，避免 OCR 误差
        rows = _text_rows(page)
        if rows:
            return rows, float(page.rect.width)
    return _ocr_rows(page)


def _text_rows(page) -> list[dict[str, Any]]:
    """从 PDF 文字层取行，用字号当行高，比例判断比边框高度可靠。"""
    rows: list[dict[str, Any]] = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            spans = [s for s in line["spans"] if s["text"].strip()]
            if not spans:
                continue
            x0, y0, x1, y1 = line["bbox"]
            rows.append(
                {
                    "x0": float(x0),
                    "x1": float(x1),
                    "y0": float(y0),
                    "y1": float(y1),
                    "h": float(max(s["size"] for s in spans)),
                    "score": 1.0,
                    "text": "".join(s["text"] for s in spans).strip(),
                }
            )
    rows.sort(key=lambda r: r["y0"])
    return rows


def _ocr_rows(page) -> tuple[list[dict[str, Any]], float]:
    """扫描页：按 200dpi 渲染后识别。"""
    pix = page.get_pixmap(dpi=200)
    pil = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    img = np.array(pil)[:, :, ::-1]  # RGB -> BGR
    h, w = img.shape[:2]
    if max(h, w) < 1400:
        scale = 1400 / max(h, w)
        img = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC)
    raw, _ = _engine()(img)
    return _rows(raw or []), float(img.shape[1])


def _preview(page, width: int = 420) -> str:
    """生成首页预览的 JPEG data URI，避免把整份 PDF 发回前端。"""
    pix = page.get_pixmap(dpi=96)
    pil = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    if pil.width > width:
        pil = pil.resize((width, int(pil.height * width / pil.width)), Image.LANCZOS)
    buf = io.BytesIO()
    pil.save(buf, format="JPEG", quality=72)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


def _rows(raw: list) -> list[dict[str, Any]]:
    """把 OCR 的框聚成视觉行，并把同一行被切碎的文字拼回去。"""
    items = []
    for box, text, score in raw:
        text = (text or "").strip()
        if not text:
            continue
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]
        x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
        items.append(
            {
                "x0": x0,
                "x1": x1,
                "y0": y0,
                "y1": y1,
                "h": max(y1 - y0, 1.0),
                "score": float(score),
                "text": text,
            }
        )

    items.sort(key=lambda it: (it["y0"] + it["y1"]) / 2)
    groups: list[dict[str, Any]] = []
    for it in items:
        cy = (it["y0"] + it["y1"]) / 2
        target = next(
            (g for g in groups if abs(cy - g["cy"]) <= 0.42 * min(it["h"], g["h"])),
            None,
        )
        if target is None:
            groups.append({"cy": cy, "h": it["h"], "items": [it]})
        else:
            target["items"].append(it)
            target["cy"] = statistics.mean(
                (i["y0"] + i["y1"]) / 2 for i in target["items"]
            )

    rows = []
    for group in groups:
        parts = sorted(group["items"], key=lambda it: it["x0"])
        h = max(it["h"] for it in parts)
        text = parts[0]["text"]
        for prev, cur in zip(parts, parts[1:]):
            gap = cur["x0"] - prev["x1"]
            text += (" " if gap > 1.2 * h else "") + cur["text"]
        rows.append(
            {
                "x0": min(it["x0"] for it in parts),
                "x1": max(it["x1"] for it in parts),
                "y0": min(it["y0"] for it in parts),
                "y1": max(it["y1"] for it in parts),
                "h": h,
                "score": min(it["score"] for it in parts),
                "text": text,
            }
        )

    rows.sort(key=lambda r: r["y0"])
    return rows


def _cx(r: dict[str, Any]) -> float:
    return (r["x0"] + r["x1"]) / 2


def _unit(r: dict[str, Any]) -> float:
    """估算字号：中文长行用平均字宽，短行退化为行高。

    OCR 的检测框高度受标点、上下伸出部件影响很大（二号标题可能只比三号正文
    高 20%），而平均字宽对中文整行更稳定。
    """
    text = r["text"]
    n = len(text)
    if n >= 6:
        cjk = sum(1 for ch in text if "\u4e00" <= ch <= "\u9fff")
        if cjk >= n * 0.7:
            return (r["x1"] - r["x0"]) / n
    return r["h"]


def _block(rows: list[dict[str, Any]], role: str | None = None) -> dict[str, Any]:
    block = {
        "text": _despace("".join(r["text"].strip() for r in rows)),
        "score": round(min(r["score"] for r in rows), 3),
        "_y": rows[0]["y0"],
    }
    if role:
        block["role"] = role
    return block


_CJK = re.compile(r"[\u3000-\u303f\u4e00-\u9fff\uff00-\uffef]")


def _despace(text: str) -> str:
    """去掉汉字之间被 OCR 误插的空格，英文单词内部不动。"""
    out: list[str] = []
    for i, ch in enumerate(text):
        if (
            ch == " "
            and i
            and i + 1 < len(text)
            and _CJK.match(text[i - 1])
            and _CJK.match(text[i + 1])
        ):
            continue
        out.append(ch)
    return "".join(out)


def _structure(rows: list[dict[str, Any]], img_w: int) -> list[dict[str, Any]]:
    body_h = statistics.median([r["h"] for r in rows]) or 1.0
    body_u = statistics.median([_unit(r) for r in rows]) or 1.0
    small = [r for r in rows if r["h"] <= 1.2 * body_h]
    left = min((r["x0"] for r in small), default=img_w * 0.08)
    right = max((r["x1"] for r in small), default=img_w * 0.92)
    # 版心中心比"各行中心的中位数"可靠：末尾短行会把中位数拉偏
    page_cx = (left + right) / 2

    # 明显大于正文、且居中的行 -> 发文机关标志 / 标题
    big = [
        r
        for r in rows
        if max(r["h"] / body_h, _unit(r) / body_u) >= 1.18
        and abs(_cx(r) - page_cx) <= 0.09 * img_w
    ]
    big_ids = {id(r) for r in big}
    rest = [r for r in rows if id(r) not in big_ids]

    blocks: list[dict[str, Any]] = []

    group: list[dict[str, Any]] = []
    for r in big:
        if group and (r["y0"] - group[-1]["y1"]) <= 0.9 * max(r["h"], group[-1]["h"]):
            group.append(r)
        else:
            if group:
                blocks.append(_block(group))
            group = [r]
    if group:
        blocks.append(_block(group))
    for b in blocks:
        b["role"] = "docorg" if b["text"].rstrip().endswith("文件") else "title"

    # 主送机关：标题下方、顶格、以冒号结尾
    for r in rest[:3]:
        if r["text"].strip().endswith(("：", ":")) and r["x0"] <= left + 1.6 * body_h:
            blocks.append(_block([r], "recipient"))
            rest.remove(r)
            break

    # 成文日期（末尾）与发文机关署名（日期上方、右偏）
    date_row = next(
        (r for r in rest[-3:] if _DATE.match(r["text"].replace(" ", ""))), None
    )
    if date_row is not None:
        idx = rest.index(date_row)
        sign = rest[idx - 1] if idx > 0 else None
        rest.pop(idx)
        if sign is not None and _cx(sign) > page_cx + 0.08 * img_w and len(sign["text"]) <= 40:
            rest.remove(sign)
            blocks.append(_block([sign], "signature"))
        blocks.append(_block([date_row], "date"))

    # 其余行：标题层级，或按首行缩进合并成正文段落
    cur: list[dict[str, Any]] = []
    for r in rest:
        text = r["text"].strip()
        if _H1.match(text):
            role = "heading1"
        elif _H2.match(text):
            role = "heading2"
        elif _H3.match(text):
            role = "heading3"
        else:
            role = None

        if role:
            if cur:
                blocks.append(_block(cur, "body"))
                cur = []
            blocks.append(_block([r], role))
            continue

        # 缩进不足一个字 -> 上一段的续行
        if cur and (r["x0"] - left) <= 0.9 * body_h:
            cur.append(r)
        else:
            if cur:
                blocks.append(_block(cur, "body"))
            cur = [r]
    if cur:
        blocks.append(_block(cur, "body"))

    blocks.sort(key=lambda b: b["_y"])
    for b in blocks:
        b.pop("_y", None)
        b.setdefault("role", "body")
    return blocks
