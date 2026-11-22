"""成熟开源库优先的解析实现。

每个函数签名与 :mod:`filegrep.extract` 里的手写版本一致；``install()`` 把可用的
库实现覆盖到注册表上。库不可用或解析失败时，自动回退到手写实现，再回退到
markitdown，最后回退到原始字符串扫描——保证任何文件都有结果。

用到的开源项目
--------------
* python-docx / openpyxl / python-pptx —— Office 三件套（保留段落、单元格地址、幻灯片结构）
* PyMuPDF > pdfplumber > pypdf > poppler pdftotext —— PDF 文本层
* py7zr / 7zz —— 7z 压缩包
* rarfile / 7zz —— rar 压缩包
* extract-msg —— Outlook .msg
* msoffcrypto-tool —— 加密 Office 解密（空密码/常见弱口令）
* markitdown —— 微软官方万能格式转换，作为未知格式兜底
* charset-normalizer —— 编码探测
"""

from __future__ import annotations

import contextlib
import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from typing import Any, Callable, Dict, List, Optional, Tuple

from . import extract as E
from .extract import ExtractContext, ExtractError, Source, decode_bytes
from .model import Document, TextBuilder

__all__ = ["install", "BACKENDS", "available_backends"]

# 记录每种格式最终使用的后端，供 GUI 展示
BACKENDS: Dict[str, str] = {}


def _try_import(name: str):
    try:
        return __import__(name)
    except Exception:
        return None


_docx_lib = _try_import("docx")
_openpyxl = _try_import("openpyxl")
_pptx_lib = _try_import("pptx")
_pymupdf = None
for _m in ("pymupdf", "fitz"):
    _pymupdf = _try_import(_m)
    if _pymupdf is not None and hasattr(_pymupdf, "open"):
        break
    _pymupdf = None
_pdfplumber = _try_import("pdfplumber")
_pypdf = _try_import("pypdf")
_py7zr = _try_import("py7zr")
_rarfile = _try_import("rarfile")
_extract_msg = _try_import("extract_msg")
_msoffcrypto = _try_import("msoffcrypto")
_markitdown = _try_import("markitdown")
_csn = _try_import("charset_normalizer")


def available_backends() -> Dict[str, str]:
    out = {}
    out["docx"] = "python-docx" if _docx_lib else "builtin(zip+xml)"
    out["xlsx"] = "openpyxl" if _openpyxl else "builtin(zip+xml)"
    out["pptx"] = "python-pptx" if _pptx_lib else "builtin(zip+xml)"
    if _pymupdf:
        out["pdf"] = "PyMuPDF"
    elif _pdfplumber:
        out["pdf"] = "pdfplumber"
    elif _pypdf:
        out["pdf"] = "pypdf"
    else:
        out["pdf"] = "none"
    out["xls"] = "xlrd" if _try_import("xlrd") else "none"
    out["7z"] = "py7zr" if _py7zr else ("7zz" if shutil.which("7zz") else "none")
    out["rar"] = "rarfile" if _rarfile else ("7zz" if shutil.which("7zz") else "none")
    out["msg"] = "extract-msg" if _extract_msg else "none"
    out["doc/rtf"] = "textutil" if shutil.which("textutil") else "none"
    out["mdb"] = "mdbtools" if shutil.which("mdb-tables") else "none"
    out["fallback"] = "markitdown" if _markitdown else "none"
    return out


# ---------------------------------------------------------------------------
# 加密 Office
# ---------------------------------------------------------------------------

_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _decrypt_office(src: Source, ctx: ExtractContext) -> Optional[io.BytesIO]:
    """若文件是加密的 OOXML，返回解密后的字节流。"""
    if _msoffcrypto is None:
        return None
    try:
        raw = src.read()
    except Exception:
        return None
    if not raw.startswith(_OLE_MAGIC):
        return None
    for pwd in ("", "123456", "1234", "password", "0000"):
        try:
            of = _msoffcrypto.OfficeFile(io.BytesIO(raw))
            if not of.is_encrypted():
                return None
            of.load_key(password=pwd)
            out = io.BytesIO()
            of.decrypt(out)
            out.seek(0)
            ctx.warn(f"文档已加密，使用口令 {pwd!r} 解密")
            return out
        except Exception:
            continue
    ctx.warn("文档已加密且未找到可用口令")
    return None


def _open_office(src: Source, ctx: ExtractContext):
    """返回 (open 函数所需的 file-like 或路径, 是否临时解密)。"""
    if src.path is not None:
        with open(src.path, "rb") as f:
            head = f.read(8)
        if head.startswith(_OLE_MAGIC):
            dec = _decrypt_office(src, ctx)
            if dec is not None:
                return dec
            raise ExtractError("加密/旧版 OLE 文档")
        return src.path
    data = src.data or b""
    if data.startswith(_OLE_MAGIC):
        dec = _decrypt_office(src, ctx)
        if dec is not None:
            return dec
        raise ExtractError("加密/旧版 OLE 文档")
    return io.BytesIO(data)


# ---------------------------------------------------------------------------
# Word
# ---------------------------------------------------------------------------

def docx_document(src: Source, ctx: ExtractContext) -> Document:
    if _docx_lib is None:
        return E.extract_docx(src, ctx)
    try:
        import docx  # noqa: F401
        from docx.table import Table
        from docx.text.paragraph import Paragraph
        from docx.oxml.ns import qn

        stream = _open_office(src, ctx)
        doc = docx.Document(stream)
    except Exception as e:
        ctx.warn(f"python-docx 失败({e})，回退内置解析")
        return E.extract_docx(src, ctx)

    b = TextBuilder(ctx.max_chars)
    counter = [0]
    tables = [0]

    def emit_paragraph(p, prefix: str = "") -> None:
        txt = p.text
        if txt and txt.strip():
            counter[0] += 1
            b.add(f"{prefix}段落 {counter[0]}", txt, {"para": counter[0]})

    def walk(parent, prefix: str = "") -> None:
        elm = parent.element.body if hasattr(parent, "element") and hasattr(parent.element, "body") else parent._tc
        for child in elm.iterchildren():
            if child.tag == qn("w:p"):
                emit_paragraph(Paragraph(child, parent), prefix)
            elif child.tag == qn("w:tbl"):
                tables[0] += 1
                tid = tables[0]
                tbl = Table(child, parent)
                for r, row in enumerate(tbl.rows, 1):
                    for c, cell in enumerate(row.cells, 1):
                        txt = cell.text.strip()
                        if txt:
                            b.add(f"{prefix}表格{tid} R{r}C{c}", txt,
                                  {"table": tid, "row": r, "col": c})

    walk(doc)
    # 页眉页脚
    for si, section in enumerate(getattr(doc, "sections", []), 1):
        for name, part in (("页眉", section.header), ("页脚", section.footer)):
            if part is None:
                continue
            try:
                for p in part.paragraphs:
                    if p.text and p.text.strip():
                        b.add(f"{name} {si}", p.text, {"header": name})
                for t in part.tables:
                    for row in t.rows:
                        vals = [c.text.strip() for c in row.cells if c.text and c.text.strip()]
                        if vals:
                            b.add(f"{name}表格 {si}", " | ".join(vals), {"header": name})
            except Exception:
                continue
    # 脚注/尾注（python-docx 无公开 API，走 zip 兜底）
    try:
        if isinstance(stream, str):
            zf = zipfile.ZipFile(stream)
        else:
            with contextlib.suppress(Exception):
                stream.seek(0)
            zf = zipfile.ZipFile(stream)
        with zf:
            for part, label in (("word/footnotes.xml", "脚注"), ("word/endnotes.xml", "尾注"),
                                ("word/comments.xml", "批注")):
                if part in zf.namelist():
                    out: List[Tuple[str, str, Dict[str, Any]]] = []
                    E._docx_part(zf, part, label + " ", out)
                    for lab, txt, meta in out:
                        b.add(lab, txt, meta)
    except Exception:
        pass

    if not len(b):
        b.add_raw("全文", "", {})
    t, segs = b.build()
    meta = {}
    with contextlib.suppress(Exception):
        cp = doc.core_properties
        for k in ("title", "author", "subject", "keywords", "comments", "last_modified_by"):
            v = getattr(cp, k, None)
            if v:
                meta[k] = str(v)
        if cp.created:
            meta["created"] = cp.created.isoformat()
    meta["paragraphs"] = counter[0]
    meta["tables"] = tables[0]
    BACKENDS["docx"] = "python-docx"
    return Document(src.name, "docx", t, segs, size=src.size, meta=meta)


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def xlsx_document(src: Source, ctx: ExtractContext) -> Document:
    if _openpyxl is None:
        return E.extract_xlsx(src, ctx)
    try:
        import openpyxl
        stream = _open_office(src, ctx)
        wb = openpyxl.load_workbook(stream, read_only=True, data_only=False)
    except Exception as e:
        ctx.warn(f"openpyxl 失败({e})，回退内置解析")
        return E.extract_xlsx(src, ctx)

    b = TextBuilder(ctx.max_chars)
    total_rows = 0
    sheets: List[str] = []

    def cell_text(v: Any) -> str:
        if v is None:
            return ""
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v).replace("\n", "\\n") if isinstance(v, str) and "\n" in v else str(v)

    try:
        for ws in wb.worksheets:
            sheets.append(ws.title)
            rows = 0
            for row in ws.iter_rows():
                rows += 1
                total_rows += 1
                if rows > 300_000:
                    ctx.warn(f"工作表 {ws.title} 超过 30 万行，已截断")
                    break
                pieces: List[str] = []
                cells: List[List[Any]] = []
                for cell in row:
                    val = cell_text(cell.value)
                    if not val:
                        continue
                    addr = cell.coordinate or ""
                    start = sum(len(p) + 3 for p in pieces)
                    pieces.append(f"{addr}={val}" if addr else val)
                    cells.append([addr, start, start + len(pieces[-1])])
                if not pieces:
                    continue
                rowno = row[0].row if row else rows
                b.add(f"工作表 {ws.title} · 第 {rowno} 行", " | ".join(pieces),
                      {"sheet": ws.title, "row": rowno, "cells": cells})
                if b.truncated:
                    break
            if b.truncated:
                break
    finally:
        with contextlib.suppress(Exception):
            wb.close()

    if not len(b):
        b.add_raw("全文", "", {})
    t, segs = b.build()
    BACKENDS["xlsx"] = "openpyxl"
    return Document(src.name, "xlsx", t, segs, size=src.size,
                    meta={"sheets": sheets, "rows": total_rows})


# ---------------------------------------------------------------------------
# PowerPoint
# ---------------------------------------------------------------------------

def pptx_document(src: Source, ctx: ExtractContext) -> Document:
    if _pptx_lib is None:
        return E.extract_pptx(src, ctx)
    try:
        import pptx
        stream = _open_office(src, ctx)
        prs = pptx.Presentation(stream)
    except Exception as e:
        ctx.warn(f"python-pptx 失败({e})，回退内置解析")
        return E.extract_pptx(src, ctx)

    b = TextBuilder(ctx.max_chars)
    slides = 0
    for idx, slide in enumerate(prs.slides, 1):
        slides += 1
        parts: List[str] = []
        for shape in slide.shapes:
            try:
                if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                    txt = shape.text_frame.text
                    if txt and txt.strip():
                        parts.append(txt.strip())
                if getattr(shape, "has_table", False):
                    try:
                        tbl = shape.table
                        for row in tbl.rows:
                            vals = [c.text.strip() for c in row.cells]
                            if any(vals):
                                parts.append(" | ".join(vals))
                    except Exception:
                        pass
            except Exception:
                continue
        if parts:
            b.add(f"幻灯片 {idx}", "\n".join(parts), {"slide": idx})
        # 备注
        try:
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text
                if notes and notes.strip():
                    b.add(f"幻灯片 {idx} · 备注", notes.strip(), {"slide": idx, "notes": True})
        except Exception:
            pass
    if not len(b):
        b.add_raw("全文", "", {})
    t, segs = b.build()
    BACKENDS["pptx"] = "python-pptx"
    return Document(src.name, "pptx", t, segs, size=src.size, meta={"slides": slides})


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

def pdf_document(src: Source, ctx: ExtractContext) -> Document:
    if _pymupdf is not None:
        try:
            return _pdf_pymupdf(src, ctx)
        except Exception as e:
            ctx.warn(f"PyMuPDF 失败({e})")
    if _pdfplumber is not None:
        try:
            return _pdf_plumber(src, ctx)
        except Exception as e:
            ctx.warn(f"pdfplumber 失败({e})")
    if shutil.which("pdftotext"):
        try:
            return _pdf_poppler(src, ctx)
        except Exception as e:
            ctx.warn(f"pdftotext 失败({e})")
    return E.extract_pdf(src, ctx)


def _pdf_pymupdf(src: Source, ctx: ExtractContext) -> Document:
    pymupdf = _pymupdf
    data = src.read() if src.data is not None else None
    doc = pymupdf.open(stream=data, filetype="pdf") if data is not None else pymupdf.open(src.path)
    b = TextBuilder(ctx.max_chars)
    pages = 0
    empty = 0
    try:
        if getattr(doc, "needs_pass", False):
            for pwd in ("", "123456", "password"):
                if doc.authenticate(pwd):
                    ctx.warn(f"PDF 已加密，使用口令 {pwd!r}")
                    break
        for i, page in enumerate(doc, 1):
            pages += 1
            try:
                txt = page.get_text("text") or ""
            except Exception as e:
                txt = ""
                ctx.warn(f"第 {i} 页提取失败: {e}")
            if not txt.strip():
                empty += 1
            b.add(f"第 {i} 页", txt, {"page": i})
            if b.truncated:
                break
        meta: Dict[str, Any] = {"pages": pages, "empty_pages": empty}
        with contextlib.suppress(Exception):
            info = doc.metadata or {}
            for k in ("title", "author", "subject", "creator", "producer", "creationDate"):
                if info.get(k):
                    meta[k] = str(info[k])
        if pages and empty == pages:
            ctx.warn("PDF 无可提取文本（扫描件，需要 OCR）")
    finally:
        with contextlib.suppress(Exception):
            doc.close()
    t, segs = b.build()
    BACKENDS["pdf"] = "PyMuPDF"
    return Document(src.name, "pdf", t, segs, size=src.size, meta=meta)


def _pdf_plumber(src: Source, ctx: ExtractContext) -> Document:
    import pdfplumber
    stream = src.open()
    b = TextBuilder(ctx.max_chars)
    pages = 0
    try:
        with pdfplumber.open(stream) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                pages += 1
                try:
                    txt = page.extract_text() or ""
                except Exception:
                    txt = ""
                b.add(f"第 {i} 页", txt, {"page": i})
                if b.truncated:
                    break
    finally:
        with contextlib.suppress(Exception):
            stream.close()
    t, segs = b.build()
    BACKENDS["pdf"] = "pdfplumber"
    return Document(src.name, "pdf", t, segs, size=src.size, meta={"pages": pages})


def _pdf_poppler(src: Source, ctx: ExtractContext) -> Document:
    with src.materialized() as path:
        out = subprocess.run(["pdftotext", "-enc", "UTF-8", "-layout", path, "-"],
                             capture_output=True, timeout=ctx.timeout)
    if out.returncode != 0:
        raise ExtractError((out.stderr or b"").decode("utf-8", "replace")[:200])
    text = out.stdout.decode("utf-8", "replace")
    pages = text.split("\f")
    b = TextBuilder(ctx.max_chars)
    for i, page in enumerate(pages, 1):
        if i == len(pages) and not page.strip():
            continue
        b.add(f"第 {i} 页", page, {"page": i})
    t, segs = b.build()
    BACKENDS["pdf"] = "pdftotext"
    return Document(src.name, "pdf", t, segs, size=src.size, meta={"pages": len(pages)})


# ---------------------------------------------------------------------------
# 压缩包（7z / rar；zip 仍走 extract 里的实现）
# ---------------------------------------------------------------------------

_RECURSABLE = None


def _recursable(suffix: str) -> bool:
    global _RECURSABLE
    if _RECURSABLE is None:
        _RECURSABLE = set(E.supported_suffixes())
    return suffix in _RECURSABLE


def _archive_members_via_7zz(src: Source, ctx: ExtractContext, b: TextBuilder) -> int:
    exe = shutil.which("7zz") or shutil.which("7z")
    if not exe:
        raise ExtractError("需要 7zz 或 py7zr")
    with src.materialized() as path:
        listing = subprocess.run([exe, "l", "-ba", "-slt", path], capture_output=True, timeout=ctx.timeout)
        names: List[str] = []
        for line in listing.stdout.decode("utf-8", "replace").splitlines():
            if line.startswith("Path = "):
                p = line[7:].strip()
                if p and p != path:
                    names.append(p)
        count = 0
        for name in names[: ctx.max_archive_members]:
            if not _recursable(os.path.splitext(name)[1].lower()):
                continue
            try:
                out = subprocess.run([exe, "x", "-so", path, name], capture_output=True, timeout=ctx.timeout)
            except subprocess.TimeoutExpired:
                ctx.warn(f"{name}: 解压超时")
                continue
            if out.returncode != 0 or not out.stdout:
                continue
            sub_ctx = ExtractContext(**{**ctx.__dict__, "warnings": []})
            sub = E.extract_source(Source(f"{src.name}::{name}", data=out.stdout), sub_ctx)
            if sub.text.strip():
                b.add(f"压缩包内 {name}", sub.text, {"member": name})
                count += 1
        return count


def archive_document(src: Source, ctx: ExtractContext) -> Document:
    suffix = src.suffix
    if ctx.archive_depth <= 0:
        return Document(src.name, suffix.lstrip("."), "", [], size=src.size,
                        warning="未启用「压缩包内部搜索」，已跳过")
    b = TextBuilder(ctx.max_chars)
    members = 0
    if suffix == ".7z" and _py7zr is not None and src.size < 200 * 1024 * 1024:
        try:
            import py7zr
            stream = src.open()
            with py7zr.SevenZipFile(stream, mode="r") as z:
                names = [n for n in z.getnames() if _recursable(os.path.splitext(n)[1].lower())]
                names = names[: ctx.max_archive_members]
                if names:
                    got = z.read(names) or {}
                    for name, bio in got.items():
                        data = bio.read() if hasattr(bio, "read") else bio
                        sub_ctx = ExtractContext(**{**ctx.__dict__, "warnings": []})
                        sub = E.extract_source(Source(f"{src.name}::{name}", data=data), sub_ctx)
                        if sub.text.strip():
                            b.add(f"压缩包内 {name}", sub.text, {"member": name})
                            members += 1
            BACKENDS["7z"] = "py7zr"
        except Exception as e:
            ctx.warn(f"py7zr 失败({e})，尝试 7zz")
    if members == 0 and suffix in (".7z", ".rar", ".xz", ".tgz", ".bz2", ".tar"):
        try:
            members = _archive_members_via_7zz(src, ctx, b)
            BACKENDS["7z" if suffix == ".7z" else "rar"] = "7zz"
        except Exception as e:
            raise ExtractError(f"无法解压 {suffix}: {e}")
    if not len(b):
        b.add_raw("全文", "", {})
    t, segs = b.build()
    return Document(src.name, suffix.lstrip("."), t, segs, size=src.size,
                    meta={"members": members})


# ---------------------------------------------------------------------------
# Outlook .msg
# ---------------------------------------------------------------------------

def msg_document(src: Source, ctx: ExtractContext) -> Document:
    if _extract_msg is None:
        return E.extract_eml(src, ctx)
    try:
        import extract_msg
        with src.materialized() as path:
            msg = extract_msg.Message(path)
            b = TextBuilder(ctx.max_chars)
            headers = []
            for label, val in (("From", getattr(msg, "sender", None)),
                               ("To", getattr(msg, "to", None)),
                               ("Cc", getattr(msg, "cc", None)),
                               ("Subject", getattr(msg, "subject", None)),
                               ("Date", getattr(msg, "date", None))):
                if val:
                    headers.append(f"{label}: {val}")
            if headers:
                b.add("邮件头", "\n".join(headers), {"header": True})
            body = getattr(msg, "body", None) or ""
            if body.strip():
                b.add("正文", body, {})
            for i, att in enumerate(getattr(msg, "attachments", []) or [], 1):
                name = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or f"附件{i}"
                try:
                    data = att.data
                except Exception:
                    data = None
                if isinstance(data, bytes) and data:
                    sub_ctx = ExtractContext(**{**ctx.__dict__, "warnings": []})
                    sub = E.extract_source(Source(f"{src.name}::{name}", data=data), sub_ctx)
                    if sub.text.strip():
                        b.add(f"附件 {name}", sub.text, {"attachment": name})
            with contextlib.suppress(Exception):
                msg.close()
            t, segs = b.build()
            BACKENDS["msg"] = "extract-msg"
            return Document(src.name, "msg", t, segs, size=src.size,
                            meta={"subject": getattr(msg, "subject", "") or ""})
    except Exception as e:
        ctx.warn(f"extract-msg 失败({e})，回退原始扫描")
        return E.extract_eml(src, ctx)


# ---------------------------------------------------------------------------
# markitdown 兜底
# ---------------------------------------------------------------------------

_md_instance = None

# markitdown 真正擅长的格式；其它二进制直接走字符串扫描，避免 magika 空转
_MD_SUFFIXES = {
    ".pdf", ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".rtf", ".odt", ".ods",
    ".odp", ".html", ".htm", ".xhtml", ".csv", ".tsv", ".json", ".jsonl", ".xml", ".epub",
    ".msg", ".ipynb", ".zip", ".txt", ".md", ".yaml", ".yml", ".eml", ".mht", ".mhtml",
}
_MD_MAX_SIZE = 32 * 1024 * 1024


def markitdown_document(src: Source, ctx: ExtractContext) -> Optional[Document]:
    global _md_instance
    if _markitdown is None:
        return None
    if src.size > _MD_MAX_SIZE or (src.suffix and src.suffix not in _MD_SUFFIXES):
        return None
    try:
        if _md_instance is None:
            from markitdown import MarkItDown
            _md_instance = MarkItDown(enable_plugins=False)
    except Exception:
        return None
    try:
        with src.materialized() as path:
            result = _md_instance.convert(path)
        text = getattr(result, "text_content", None) or ""
    except Exception:
        return None
    if not text.strip():
        return None
    b = TextBuilder(ctx.max_chars)
    b.add_raw("MarkItDown 转换", text, {"via": "markitdown"})
    t, segs = b.build()
    BACKENDS["markitdown"] = "markitdown"
    return Document(src.name, "text", t, segs, size=src.size, meta={"via": "markitdown"})


# ---------------------------------------------------------------------------
# 安装
# ---------------------------------------------------------------------------

def install() -> None:
    """把库实现覆盖到 extract 的注册表（仅覆盖可用的）。"""
    mapping: List[Tuple[Tuple[str, ...], Callable[[Source, ExtractContext], Document]]] = [
        ((".docx", ".docm", ".dotx", ".dotm"), docx_document),
        ((".xlsx", ".xlsm", ".xltx", ".xltm"), xlsx_document),
        ((".pptx", ".pptm", ".ppsx", ".potx"), pptx_document),
        ((".pdf",), pdf_document),
        ((".7z", ".rar", ".xz", ".tgz", ".bz2", ".tar"), archive_document),
        ((".msg",), msg_document),
    ]
    for suffixes, fn in mapping:
        for s in suffixes:
            E._REGISTRY[s] = fn
            E._SUFFIX_KIND.setdefault(s, s.lstrip("."))
    # 加密 Office 的 .docx/.xlsx/.pptx 也能被识别
    for s in (".docx", ".xlsx", ".pptx"):
        E._SUFFIX_KIND[s] = s.lstrip(".")
    E._FALLBACK_CONVERTER = markitdown_document  # 由 extract_source 调用


def enhanced_decode(raw: bytes) -> Tuple[str, str]:
    """charset-normalizer 辅助的编码判定（仅在 utf-8 失败时启用）。"""
    try:
        return raw.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        pass
    if _csn is not None:
        with contextlib.suppress(Exception):
            best = _csn.from_bytes(raw[:200_000]).best()
            if best is not None:
                enc = best.encoding
                with contextlib.suppress(Exception):
                    return raw.decode(enc), enc
    return decode_bytes(raw)
