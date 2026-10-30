"""图片 / PDF -> 公文格式 Word 的本地 Web 应用。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, Response
from pydantic import BaseModel, Field

import docx_writer
import ocr_engine

BASE = Path(__file__).resolve().parent
MAX_BYTES = 50 * 1024 * 1024
MAX_FILES = 10

app = FastAPI(title="公文 OCR")


class BlockIn(BaseModel):
    role: str = "body"
    text: str = ""


class DocxIn(BaseModel):
    blocks: list[BlockIn] = Field(default_factory=list)
    body_font: str = "仿宋"
    title_font: str = "黑体"


@app.get("/")
def index() -> FileResponse:
    return FileResponse(BASE / "static" / "index.html")


@app.post("/api/ocr")
async def api_ocr(files: list[UploadFile] = File(...)) -> dict:
    if len(files) > MAX_FILES:
        raise HTTPException(400, f"一次最多上传 {MAX_FILES} 个文件")

    uploads = [(f.filename or "", await f.read()) for f in files]
    if not uploads:
        raise HTTPException(400, "请选择文件")
    if sum(len(d) for _, d in uploads) > MAX_BYTES:
        raise HTTPException(413, "文件总大小超过 50MB，请压缩后再试")

    is_pdf = [
        d[:5] == b"%PDF-" or name.lower().endswith(".pdf") for name, d in uploads
    ]
    if any(is_pdf) and len(uploads) > 1:
        raise HTTPException(400, "PDF 请单独上传，图片可以多选")

    try:
        if is_pdf[0]:
            return ocr_engine.analyze_pdf(uploads[0][1])
        return ocr_engine.analyze_many([d for _, d in uploads])
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc


@app.post("/api/docx")
def api_docx(payload: DocxIn) -> Response:
    blocks = [
        {"role": b.role, "text": b.text.strip()}
        for b in payload.blocks
        if b.role in ("blank", "pagebreak") or b.text.strip()
    ]
    while blocks and blocks[-1]["role"] == "pagebreak":
        blocks.pop()  # 末尾的分页符会多出一个空白页
    if not blocks:
        raise HTTPException(400, "没有可导出的内容")

    data = docx_writer.build_docx(
        blocks, body_font=payload.body_font, title_font=payload.title_font
    )
    name = "公文_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".docx"
    return Response(
        content=data,
        media_type=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{quote(name)}"
        },
    )
