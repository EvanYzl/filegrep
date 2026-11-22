"""文件解析器：把各种格式抽取成「带定位信息的纯文本」。

统一约定
--------
* 所有抽取器返回 :class:`Document`，其中 ``text`` 是全文，``segments`` 记录
  每段文本的来源（页码 / 工作表+行 / 数据库表+行 / 压缩包成员 ...）。
* 抽取器签名 ``fn(src: Source, ctx: ExtractContext) -> Document``。
* 解析失败不抛到外层：由 :func:`extract_source` 统一捕获，退化到原始字符串扫描
  或返回 error 字段，保证批量扫描永远不会中断。

零第三方依赖：docx/xlsx/pptx/odt/epub 用 zipfile + ElementTree 自行解析。
可选依赖：pypdf（PDF）、xlrd（老式 .xls）、textutil（macOS 自带，处理 .doc/.rtf）。
"""

from __future__ import annotations

import contextlib
import io
import os
import re
import shutil
import sqlite3
import subprocess
import tempfile
import time
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from email import message_from_binary_file
from email import policy as email_policy
from html.parser import HTMLParser
from typing import Any, Callable, Dict, Iterator, List, Optional, Tuple
import xml.etree.ElementTree as ET

from .model import Document, Segment, TextBuilder

__all__ = [
    "ExtractContext", "Source", "ExtractError", "extract_path", "extract_source",
    "SUFFIX_KINDS", "kind_of", "supported_suffixes", "decode_bytes",
]


class ExtractError(Exception):
    """解析失败（可捕获，会转成 Document.error）。"""


# ---------------------------------------------------------------------------
# 上下文
# ---------------------------------------------------------------------------

@dataclass
class ExtractContext:
    max_chars: int = 40_000_000          # 单文件抽取文本上限（字符）
    sqlite_max_rows: int = 200_000       # 单表最多读取行数
    archive_depth: int = 1               # 递归解压深度（1=只解一层）
    max_archive_members: int = 2000
    raw_strings: bool = False            # 对未知/二进制文件做 strings 扫描
    zip_passwords: List[str] = field(default_factory=list)   # 加密压缩包口令
    use_textutil: bool = True            # macOS 上用 textutil 处理 doc/rtf
    timeout: float = 120.0
    warnings: List[str] = field(default_factory=list)

    def warn(self, msg: str) -> None:
        if len(self.warnings) < 50:
            self.warnings.append(msg)


# ---------------------------------------------------------------------------
# 输入源：磁盘文件或内存字节（压缩包成员、数据库 BLOB）
# ---------------------------------------------------------------------------

class Source:
    __slots__ = ("name", "path", "data", "size", "_suffix")

    def __init__(self, name: str, *, path: Optional[str] = None,
                 data: Optional[bytes] = None, size: Optional[int] = None):
        if path is None and data is None:
            # 允许 Source("/real/path") 简写：磁盘上的名字即路径
            path = name
        self.name = name
        self.path = path
        self.data = data
        if size is None:
            if path is not None:
                try:
                    size = os.path.getsize(path)
                except OSError:
                    size = 0
            else:
                size = len(data or b"")
        self.size = size
        base = name.rsplit("/", 1)[-1]
        self._suffix = os.path.splitext(base)[1].lower()

    @property
    def suffix(self) -> str:
        return self._suffix

    @property
    def basename(self) -> str:
        return os.path.basename(self.path) if self.path else self.name.rsplit("/", 1)[-1]

    def open(self):
        if self.path is not None:
            return open(self.path, "rb")
        return io.BytesIO(self.data or b"")

    def read(self, n: int = -1) -> bytes:
        if self.data is not None:
            return self.data if n < 0 else self.data[:n]
        with open(self.path, "rb") as f:
            return f.read(n)

    @contextlib.contextmanager
    def materialized(self) -> Iterator[str]:
        """保证有真实磁盘路径（数据库、外部工具需要）。"""
        if self.path is not None:
            yield self.path
            return
        fd, tmp = tempfile.mkstemp(prefix="filegrep_", suffix=self.suffix or ".bin")
        try:
            with os.fdopen(fd, "wb") as f:
                f.write(self.data or b"")
            yield tmp
        finally:
            with contextlib.suppress(OSError):
                os.unlink(tmp)

    def __repr__(self) -> str:  # pragma: no cover
        return f"<Source {self.name} suffix={self.suffix} size={self.size}>"


# ---------------------------------------------------------------------------
# 编码判定
# ---------------------------------------------------------------------------

_BOMS = (
    (b"\xef\xbb\xbf", "utf-8-sig"),
    (b"\xff\xfe\x00\x00", "utf-32-le"),
    (b"\x00\x00\xfe\xff", "utf-32-be"),
    (b"\xff\xfe", "utf-16-le"),
    (b"\xfe\xff", "utf-16-be"),
)

_TEXT_CANDIDATES = ("utf-8", "gb18030", "big5", "shift_jis", "euc-kr", "cp1252")


def _text_score(s: str) -> float:
    """粗估解码质量：可读字符占比（0~1）。"""
    if not s:
        return 0.0
    good = bad = 0.0
    for ch in s[:20000]:
        o = ord(ch)
        if o == 0xFFFD or 0xE000 <= o <= 0xF8FF:
            bad += 2.0
        elif ch in "\n\r\t":
            good += 1.0
        elif 0x20 <= o <= 0x7E:
            good += 1.0
        elif 0x4E00 <= o <= 0x9FFF:
            good += 1.0
        elif 0x3000 <= o <= 0x303F or 0xFF00 <= o <= 0xFFEF:
            good += 1.0
        elif 0x2000 <= o <= 0x2BFF:
            good += 0.8
        elif o < 0x20 or 0x7F <= o <= 0x9F:
            bad += 2.0
        elif 0xAC00 <= o <= 0xD7AF or 0x3040 <= o <= 0x30FF:
            good += 1.0
        else:
            good += 0.4
    return good / (good + bad + 1.0)


def decode_bytes(raw: bytes) -> Tuple[str, str]:
    """尽力解码，返回 (文本, 编码名)。中文环境优先 GB18030。"""
    if not raw:
        return "", "utf-8"
    for bom, enc in _BOMS:
        if raw.startswith(bom):
            with contextlib.suppress(UnicodeDecodeError):
                return raw.decode(enc), enc
    try:
        return raw.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        pass
    head = raw[:8192]
    if head.count(b"\x00") > len(head) * 0.2:
        for enc in ("utf-16-le", "utf-16-be"):
            with contextlib.suppress(UnicodeDecodeError):
                s = raw.decode(enc)
                if _text_score(s) > 0.6:
                    return s, enc
    best: Optional[Tuple[float, str, str]] = None
    for enc in _TEXT_CANDIDATES:
        try:
            s = raw.decode(enc)
        except UnicodeDecodeError:
            s = raw.decode(enc, errors="replace")
        score = _text_score(s) - (0.15 if "\ufffd" in s else 0.0)
        if best is None or score > best[0]:
            best = (score, s, enc)
    if best is None:
        return raw.decode("utf-8", errors="replace"), "utf-8/replace"
    return best[1], best[2]


# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------

def _local(tag: str) -> str:
    """去掉 XML 命名空间，返回本地名。"""
    return tag.rpartition("}")[2] if "}" in tag else tag


def _kids(node, name: str) -> Iterator[ET.Element]:
    for child in node:
        if _local(child.tag) == name:
            yield child


def _first(node, name: str) -> Optional[ET.Element]:
    for child in node:
        if _local(child.tag) == name:
            return child
    return None


def _xml(zf: zipfile.ZipFile, name: str) -> Optional[ET.Element]:
    try:
        raw = zf.read(name)
    except KeyError:
        return None
    except Exception:
        return None
    try:
        return ET.fromstring(raw)
    except ET.ParseError:
        return None


def _xml_name(zf: zipfile.ZipFile, name: str) -> Optional[bytes]:
    try:
        return zf.read(name)
    except KeyError:
        return None


def _rels(zf: zipfile.ZipFile, part: str) -> Dict[str, str]:
    """读取 OOXML 关系文件，返回 rId -> 目标路径（已规范化到包内绝对路径）。"""
    d = os.path.dirname(part)
    rel_name = (d + "/_rels/" + os.path.basename(part) + ".rels") if d else "_rels/" + os.path.basename(part) + ".rels"
    root = _xml(zf, rel_name)
    if root is None:
        return {}
    out: Dict[str, str] = {}
    for rel in root:
        if _local(rel.tag) != "Relationship":
            continue
        rid = rel.get("Id") or ""
        target = (rel.get("Target") or "").replace("\\", "/")
        if not target or rel.get("TargetMode") == "External":
            continue
        if target.startswith("/"):
            full = target.lstrip("/")
        else:
            full = os.path.normpath(os.path.join(d, target)).replace("\\", "/")
        out[rid] = full
    return out


def _fmt_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f}{unit}" if unit == "B" else f"{n:.1f}{unit}"
        n /= 1024.0
    return f"{n}B"


# ---------------------------------------------------------------------------
# 注册表
# ---------------------------------------------------------------------------

Extractor = Callable[[Source, ExtractContext], Document]

# 由 :mod:`filegrep.parsers` 注入（markitdown 万能转换兜底）
_FALLBACK_CONVERTER: Optional[Callable[[Source, ExtractContext], Optional[Document]]] = None

_REGISTRY: Dict[str, Extractor] = {}
_SUFFIX_KIND: Dict[str, str] = {}


def register(kind: str, *suffixes: str):
    def deco(fn: Extractor) -> Extractor:
        for s in suffixes:
            _REGISTRY[s.lower()] = fn
            _SUFFIX_KIND[s.lower()] = kind
        return fn
    return deco


def kind_of(path_or_suffix: str) -> str:
    s = path_or_suffix.lower()
    if not s.startswith("."):
        s = os.path.splitext(s)[1].lower()
    return _SUFFIX_KIND.get(s, "binary")


def supported_suffixes() -> List[str]:
    return sorted(_REGISTRY)


SUFFIX_KINDS = _SUFFIX_KIND   # 向后兼容别名


# ---------------------------------------------------------------------------
# 纯文本族
# ---------------------------------------------------------------------------

TEXT_SUFFIXES = (
    ".txt", ".text", ".log", ".md", ".markdown", ".rst", ".nfo", ".readme",
    ".json", ".jsonl", ".ndjson", ".xml", ".yaml", ".yml", ".toml", ".ini",
    ".conf", ".cfg", ".properties", ".env", ".list", ".lst",
    ".html", ".htm", ".xhtml", ".css", ".scss", ".less",
    ".js", ".mjs", ".ts", ".jsx", ".tsx", ".vue", ".py", ".java", ".kt",
    ".c", ".h", ".cpp", ".hpp", ".cc", ".cs", ".go", ".rs", ".rb", ".php",
    ".pl", ".lua", ".sh", ".bash", ".zsh", ".bat", ".cmd", ".ps1", ".psm1",
    ".sql", ".reg", ".vbs", ".ahk", ".asm", ".s", ".vb", ".swift", ".m",
    ".diff", ".patch", ".gitignore", ".dockerfile", ".makefile", ".mk",
    ".srt", ".ass", ".vtt", ".lrc", ".eml", ".mbox", ".vcf", ".ics",
    ".svg", ".plist", ".strings", ".po", ".pot", ".tmx", ".resx",
)


@register("text", *TEXT_SUFFIXES)
def extract_text(src: Source, ctx: ExtractContext) -> Document:
    raw = src.read(min(src.size, ctx.max_chars * 4) if ctx.max_chars else -1)
    text, enc = decode_bytes(raw)
    if ctx.max_chars and len(text) > ctx.max_chars:
        text = text[: ctx.max_chars]
        ctx.warn("文本超长已截断")
    b = TextBuilder()
    b.add_raw("全文", text, {"encoding": enc})
    t, segs = b.build()
    return Document(src.name, "text", t, segs, size=src.size, meta={"encoding": enc})


@register("code", ".pyc", ".class", ".o", ".so", ".dylib", ".dll", ".exe", ".bin")
def extract_binary_strings(src: Source, ctx: ExtractContext) -> Document:
    """可执行/目标文件：抽取字符串常量。"""
    return _strings_document(src, ctx, kind="binary")


@register("csv", ".csv", ".tsv", ".tab", ".psv")
def extract_csv(src: Source, ctx: ExtractContext) -> Document:
    raw = src.read()
    text, enc = decode_bytes(raw)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    sep = "\t" if src.suffix in (".tsv", ".tab") else ("," if src.suffix == ".csv" else "|")
    b = TextBuilder()
    if len(lines) > 200_000:
        b.add_raw("全文", text, {"encoding": enc, "rows": len(lines)})
    else:
        for i, line in enumerate(lines, 1):
            if line.strip():
                b.add(f"第 {i} 行", line, {"row": i, "encoding": enc, "sep": sep})
    t, segs = b.build()
    return Document(src.name, "csv", t, segs, size=src.size,
                    meta={"encoding": enc, "rows": len(lines), "columns": _csv_header(lines, sep)})


def _csv_header(lines: List[str], sep: str) -> List[str]:
    for line in lines[:5]:
        if line.strip():
            for cand in (sep, ",", "\t", "|", ";"):
                parts = [p.strip().strip('"') for p in line.split(cand)]
                if len(parts) > 1:
                    return parts[:64]
    return []


# ---------------------------------------------------------------------------
# WordprocessingML (.docx/.docm/.dotx)
# ---------------------------------------------------------------------------

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w_para_text(p: ET.Element) -> str:
    out: List[str] = []
    for el in p.iter():
        ln = _local(el.tag)
        if ln == "t":
            out.append(el.text or "")
        elif ln == "tab":
            out.append("\t")
        elif ln in ("br", "cr"):
            out.append("\n")
        elif ln == "noBreakHyphen":
            out.append("-")
        elif ln == "softHyphen":
            out.append("")
    return "".join(out)


def _walk_wml(node: ET.Element, out: List[Tuple[str, str, Dict[str, Any]]],
              prefix: str, counter: List[int], tbl: List[int]) -> None:
    for child in node:
        ln = _local(child.tag)
        if ln == "p":
            txt = _w_para_text(child)
            if txt.strip():
                counter[0] += 1
                out.append((f"{prefix}段落 {counter[0]}", txt, {"para": counter[0]}))
        elif ln == "tbl":
            tbl[0] += 1
            tid = tbl[0]
            for r, tr in enumerate(_kids(child, "tr"), 1):
                for c, tc in enumerate(_kids(tr, "tc"), 1):
                    _walk_wml(tc, out, f"{prefix}表格{tid} R{r}C{c} ", counter, tbl)
        elif ln in ("body", "sdt", "sdtContent", "txbxContent", "hdr", "ftr",
                    "footnote", "endnote", "comment", "document", "tc", "txbx"):
            _walk_wml(child, out, prefix, counter, tbl)


def _docx_part(zf: zipfile.ZipFile, part: str, label_prefix: str,
               out: List[Tuple[str, str, Dict[str, Any]]]) -> None:
    root = _xml(zf, part)
    if root is None:
        return
    _walk_wml(root, out, label_prefix, [0], [0])


@register("docx", ".docx", ".docm", ".dotx", ".dotm", ".dot")
def extract_docx(src: Source, ctx: ExtractContext) -> Document:
    # .dot 走 textutil 分支
    if src.suffix == ".dot":
        return _textutil_document(src, ctx, "doc")
    if src.path is not None:
        zf_cm = zipfile.ZipFile(src.path)
    else:
        zf_cm = zipfile.ZipFile(io.BytesIO(src.data or b""))
    with zf_cm as zf:
        names = set(zf.namelist())
        if "word/document.xml" not in names:
            raise ExtractError("不是有效的 docx（缺少 word/document.xml）")
        out: List[Tuple[str, str, Dict[str, Any]]] = []
        _docx_part(zf, "word/document.xml", "", out)
        for n in sorted(names):
            if re.match(r"word/(header|footer)\d*\.xml$", n):
                _docx_part(zf, n, ("页眉 " if "header" in n else "页脚 "), out)
            elif n in ("word/footnotes.xml",):
                _docx_part(zf, n, "脚注 ", out)
            elif n in ("word/endnotes.xml",):
                _docx_part(zf, n, "尾注 ", out)
            elif n in ("word/comments.xml",):
                _docx_part(zf, n, "批注 ", out)
        b = TextBuilder()
        for label, text, meta in out:
            b.add(label, text, meta)
        if not len(b):
            b.add_raw("全文", "", {})
        t, segs = b.build()
        meta = _docx_core(zf)
        meta["paragraphs"] = len(out)
        return Document(src.name, "docx", t, segs, size=src.size, meta=meta)


def _docx_core(zf: zipfile.ZipFile) -> Dict[str, Any]:
    root = _xml(zf, "docProps/core.xml")
    if root is None:
        return {}
    out: Dict[str, Any] = {}
    for el in root:
        ln = _local(el.tag)
        if ln in ("title", "creator", "subject", "description", "lastModifiedBy",
                  "created", "modified", "keywords", "category"):
            if el.text:
                out[ln] = el.text.strip()
    return out


# ---------------------------------------------------------------------------
# SpreadsheetML (.xlsx/.xlsm/.xltx)
# ---------------------------------------------------------------------------

_DATE_NUMFMT = set(range(14, 23)) | {45, 46, 47}


def _xlsx_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    root = _xml(zf, "xl/sharedStrings.xml")
    if root is None:
        return []
    out: List[str] = []

    def collect(node: ET.Element, buf: List[str]) -> None:
        for child in node:
            ln = _local(child.tag)
            if ln == "rPh":          # 注音，跳过
                continue
            if ln == "t":
                buf.append(child.text or "")
            else:
                collect(child, buf)

    for si in root:
        if _local(si.tag) != "si":
            continue
        buf: List[str] = []
        collect(si, buf)
        out.append("".join(buf))
    return out


def _xlsx_date_styles(zf: zipfile.ZipFile) -> set:
    root = _xml(zf, "xl/styles.xml")
    if root is None:
        return set()
    custom: Dict[int, str] = {}
    for el in root.iter():
        if _local(el.tag) == "numFmt":
            with contextlib.suppress(ValueError):
                custom[int(el.get("numFmtId", "0"))] = el.get("formatCode", "")
    cellxfs = _first(root, "cellXfs")
    out = set()
    if cellxfs is None:
        return out
    for i, xf in enumerate(cellxfs):
        if _local(xf.tag) != "xf":
            continue
        try:
            nid = int(xf.get("numFmtId", "0"))
        except ValueError:
            continue
        code = custom.get(nid, "")
        if nid in _DATE_NUMFMT or (code and re.search(r"[yYdD]", code) and "[" not in code):
            out.add(i)
    return out


def _excel_serial_to_str(v: str) -> str:
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if not (1 <= f <= 2958465):
        return str(v)
    base = datetime(1899, 12, 30)
    dt = base + timedelta(days=f)
    # 浮点误差收敛到秒
    dt = dt.replace(microsecond=0)
    if abs(dt.hour) + abs(dt.minute) + abs(dt.second) == 0:
        return dt.strftime("%Y-%m-%d")
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def _xlsx_cell_value(c: ET.Element, shared: List[str], date_styles: set) -> str:
    t = c.get("t") or "n"
    if t == "inlineStr":
        is_node = _first(c, "is")
        if is_node is None:
            return ""
        buf: List[str] = []

        def collect(node: ET.Element) -> None:
            for ch in node:
                if _local(ch.tag) == "rPh":
                    continue
                if _local(ch.tag) == "t":
                    buf.append(ch.text or "")
                else:
                    collect(ch)
        collect(is_node)
        return "".join(buf)
    v = _first(c, "v")
    val = (v.text or "") if v is not None else ""
    if t == "s":
        try:
            return shared[int(val)]
        except (ValueError, IndexError):
            return val
    if t == "b":
        return "TRUE" if val.strip() in ("1", "TRUE", "true") else "FALSE"
    if t in ("str", "e", "d"):
        return val
    sty = c.get("s")
    if sty is not None and date_styles:
        with contextlib.suppress(ValueError):
            if int(sty) in date_styles:
                return _excel_serial_to_str(val)
    return val


@register("xlsx", ".xlsx", ".xlsm", ".xltx", ".xltm")
def extract_xlsx(src: Source, ctx: ExtractContext) -> Document:
    zf_cm = zipfile.ZipFile(src.path) if src.path is not None else zipfile.ZipFile(io.BytesIO(src.data or b""))
    with zf_cm as zf:
        if "xl/workbook.xml" not in zf.namelist():
            raise ExtractError("不是有效的 xlsx（缺少 xl/workbook.xml）")
        shared = _xlsx_shared_strings(zf)
        date_styles = _xlsx_date_styles(zf)
        wb = _xml(zf, "xl/workbook.xml")
        rels = _rels(zf, "xl/workbook.xml")
        sheets: List[Tuple[str, str]] = []
        if wb is not None:
            container = _first(wb, "sheets")
            if container is not None:
                for sh in container:
                    if _local(sh.tag) != "sheet":
                        continue
                    name = sh.get("name") or "Sheet"
                    rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") or sh.get("id")
                    target = rels.get(rid or "", "")
                    if not target:
                        continue
                    sheets.append((name, target))
        if not sheets:
            sheets = [(os.path.basename(n), n) for n in zf.namelist() if n.startswith("xl/worksheets/sheet")]
        b = TextBuilder(ctx.max_chars)
        total_rows = 0
        hidden: List[str] = []
        for sname, part in sheets:
            root = _xml(zf, part)
            if root is None:
                continue
            sheet_data = _first(root, "sheetData")
            if sheet_data is None:
                continue
            rows = 0
            for row in _kids(sheet_data, "row"):
                rows += 1
                total_rows += 1
                pieces: List[str] = []
                cells: List[List[Any]] = []
                for c in _kids(row, "c"):
                    addr = c.get("r") or ""
                    val = _xlsx_cell_value(c, shared, date_styles)
                    if val == "":
                        continue
                    start = sum(len(p) + 1 for p in pieces)
                    pieces.append(f"{addr}={val}" if addr else val)
                    cells.append([addr, start, start + len(pieces[-1])])
                if not pieces:
                    continue
                line = " | ".join(pieces)
                rowno = row.get("r") or str(rows)
                b.add(f"工作表 {sname} · 第 {rowno} 行", line,
                      {"sheet": sname, "row": int(rowno) if str(rowno).isdigit() else rows,
                       "cells": cells})
                if b.truncated:
                    break
            if b.truncated:
                break
        t, segs = b.build()
        return Document(src.name, "xlsx", t, segs, size=src.size,
                        meta={"sheets": [s[0] for s in sheets], "rows": total_rows,
                              "shared_strings": len(shared)})


# ---------------------------------------------------------------------------
# PresentationML (.pptx)
# ---------------------------------------------------------------------------

def _a_text(node: ET.Element) -> str:
    out: List[str] = []
    for el in node.iter():
        ln = _local(el.tag)
        if ln == "t":
            out.append(el.text or "")
        elif ln == "br":
            out.append("\n")
    return "".join(out)


@register("pptx", ".pptx", ".pptm", ".ppsx", ".potx")
def extract_pptx(src: Source, ctx: ExtractContext) -> Document:
    zf_cm = zipfile.ZipFile(src.path) if src.path is not None else zipfile.ZipFile(io.BytesIO(src.data or b""))
    with zf_cm as zf:
        names = zf.namelist()
        if "ppt/presentation.xml" not in names:
            raise ExtractError("不是有效的 pptx（缺少 ppt/presentation.xml）")
        rels = _rels(zf, "ppt/presentation.xml")
        pres = _xml(zf, "ppt/presentation.xml")
        order: List[Tuple[int, str]] = []
        if pres is not None:
            lst = _first(pres, "sldIdLst")
            if lst is not None:
                for i, sid in enumerate(lst, 1):
                    rid = sid.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") or sid.get("id")
                    part = rels.get(rid or "")
                    if part:
                        order.append((i, part))
        if not order:
            slides = [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)]
            slides.sort(key=lambda n: int(re.search(r"(\d+)", os.path.basename(n)).group(1)))
            order = [(i, n) for i, n in enumerate(slides, 1)]
        b = TextBuilder(ctx.max_chars)
        counted = 0
        for idx, part in order:
            root = _xml(zf, part)
            if root is None:
                continue
            texts: List[str] = []
            for sp in root.iter():
                if _local(sp.tag) in ("txBody", "sp"):
                    for p in _kids(sp, "p"):
                        txt = _a_text(p)
                        if txt.strip():
                            texts.append(txt)
                    if texts:
                        break
            if not texts:
                for p in root.iter():
                    if _local(p.tag) == "p":
                        txt = _a_text(p)
                        if txt.strip():
                            texts.append(txt)
            if texts:
                counted += 1
                b.add(f"幻灯片 {idx}", "\n".join(texts), {"slide": idx})
            # 备注
            nrels = _rels(zf, part)
            for rid, tgt in nrels.items():
                if "notesSlide" in tgt:
                    nroot = _xml(zf, tgt)
                    if nroot is None:
                        continue
                    nt: List[str] = []
                    for p in nroot.iter():
                        if _local(p.tag) == "p":
                            txt = _a_text(p)
                            if txt.strip():
                                nt.append(txt)
                    if nt:
                        b.add(f"幻灯片 {idx} · 备注", "\n".join(nt), {"slide": idx, "notes": True})
        t, segs = b.build()
        return Document(src.name, "pptx", t, segs, size=src.size,
                        meta={"slides": len(order), "text_slides": counted})


# ---------------------------------------------------------------------------
# OpenDocument / EPUB
# ---------------------------------------------------------------------------

@register("odf", ".odt", ".ods", ".odp", ".odg", ".fodt")
def extract_odf(src: Source, ctx: ExtractContext) -> Document:
    zf_cm = zipfile.ZipFile(src.path) if src.path is not None else zipfile.ZipFile(io.BytesIO(src.data or b""))
    with zf_cm as zf:
        root = _xml(zf, "content.xml")
        if root is None:
            raise ExtractError("不是有效的 ODF（缺少 content.xml）")
        b = TextBuilder(ctx.max_chars)
        para_no = 0

        def collect(node: ET.Element, table_ctx: Optional[Tuple[str, int, int]] = None) -> None:
            nonlocal para_no
            for child in node:
                ln = _local(child.tag)
                if ln in ("p", "h"):
                    txt = _a_text(child)
                    if txt.strip():
                        para_no += 1
                        label = f"段落 {para_no}" if table_ctx is None else \
                            f"表 {table_ctx[0]} R{table_ctx[1]}C{table_ctx[2]}"
                        b.add(label, txt, {"para": para_no})
                elif ln == "table":
                    tname = child.get("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}name") or "table"
                    for r, row in enumerate(_kids(child, "table-row"), 1):
                        for c, cell in enumerate(_kids(row, "table-cell"), 1):
                            collect(cell, (tname, r, c))
                else:
                    collect(child, table_ctx)

        collect(root)
        t, segs = b.build()
        return Document(src.name, "odf", t, segs, size=src.size, meta={})


class _HTMLText(HTMLParser):
    """HTML -> 纯文本（保留段落换行）。"""

    BLOCK = {"p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6",
             "table", "section", "article", "header", "footer", "pre", "blockquote"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: List[str] = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self._skip += 1
        elif tag in self.BLOCK:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style") and self._skip:
            self._skip -= 1
        elif tag in self.BLOCK:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self._skip:
            self.parts.append(data)

    def text(self) -> str:
        raw = "".join(self.parts)
        raw = re.sub(r"[ \t\u00a0]+", " ", raw)
        raw = re.sub(r"\n\s*\n\s*\n+", "\n\n", raw)
        return raw.strip()


def html_to_text(markup: str) -> str:
    p = _HTMLText()
    try:
        p.feed(markup)
    except Exception:
        return re.sub(r"<[^>]+>", " ", markup)
    return p.text()


@register("epub", ".epub")
def extract_epub(src: Source, ctx: ExtractContext) -> Document:
    zf_cm = zipfile.ZipFile(src.path) if src.path is not None else zipfile.ZipFile(io.BytesIO(src.data or b""))
    with zf_cm as zf:
        names = zf.namelist()
        opf = next((n for n in names if n.endswith(".opf")), None)
        order: List[str] = []
        title = ""
        if opf:
            root = _xml(zf, opf)
            if root is not None:
                base = os.path.dirname(opf)
                for el in root.iter():
                    if _local(el.tag) == "title" and el.text and not title:
                        title = el.text.strip()
                manifest = {}
                for el in root.iter():
                    if _local(el.tag) == "item":
                        manifest[el.get("id") or ""] = el.get("href") or ""
                for el in root.iter():
                    if _local(el.tag) == "itemref":
                        href = manifest.get(el.get("idref") or "", "")
                        if href:
                            order.append(os.path.normpath(os.path.join(base, href)).replace("\\", "/"))
        if not order:
            order = sorted(n for n in names if n.endswith((".xhtml", ".html", ".htm")))
        b = TextBuilder(ctx.max_chars)
        for i, part in enumerate(order, 1):
            raw = _xml_name(zf, part)
            if raw is None:
                continue
            text = html_to_text(decode_bytes(raw)[0])
            if text:
                b.add(f"章节 {i} · {os.path.basename(part)}", text, {"chapter": i, "part": part})
        t, segs = b.build()
        return Document(src.name, "epub", t, segs, size=src.size, meta={"title": title})


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

@register("pdf", ".pdf")
def extract_pdf(src: Source, ctx: ExtractContext) -> Document:
    try:
        from pypdf import PdfReader
    except ImportError:  # pragma: no cover
        raise ExtractError("需要 pypdf：pip install pypdf")
    stream = src.open()
    try:
        reader = PdfReader(stream)
        if getattr(reader, "is_encrypted", False):
            for pwd in ("", "123456", "password"):
                with contextlib.suppress(Exception):
                    if reader.decrypt(pwd):
                        break
        b = TextBuilder(ctx.max_chars)
        pages = 0
        empty = 0
        for i, page in enumerate(reader.pages, 1):
            pages += 1
            try:
                txt = page.extract_text() or ""
            except Exception as e:
                txt = ""
                ctx.warn(f"第 {i} 页解析失败: {e}")
            if not txt.strip():
                empty += 1
            b.add(f"第 {i} 页", txt, {"page": i})
            if b.truncated:
                break
        t, segs = b.build()
        meta: Dict[str, Any] = {"pages": pages, "empty_pages": empty}
        with contextlib.suppress(Exception):
            info = reader.metadata or {}
            for k in ("/Title", "/Author", "/Subject", "/Creator", "/Producer"):
                if info.get(k):
                    meta[k.lstrip("/").lower()] = str(info.get(k))
        if empty and pages and empty == pages:
            ctx.warn("PDF 无可提取文本层（可能是扫描件，需要 OCR）")
        return Document(src.name, "pdf", t, segs, size=src.size, meta=meta)
    finally:
        with contextlib.suppress(Exception):
            stream.close()


# ---------------------------------------------------------------------------
# 老式 Office / RTF：macOS textutil 或 linux 工具
# ---------------------------------------------------------------------------

def _run(cmd: List[str], timeout: float) -> bytes:
    r = subprocess.run(cmd, capture_output=True, timeout=timeout)
    if r.returncode != 0:
        raise ExtractError((r.stderr or b"").decode("utf-8", "replace")[:300] or f"{cmd[0]} 退出码 {r.returncode}")
    return r.stdout


def _textutil_document(src: Source, ctx: ExtractContext, kind: str) -> Document:
    if not ctx.use_textutil or not shutil.which("textutil"):
        raise ExtractError("需要 textutil（macOS 自带）或安装 antiword/catdoc")
    with src.materialized() as path:
        raw = _run(["/usr/bin/textutil", "-convert", "txt", "-stdout", path], ctx.timeout)
    text = decode_bytes(raw)[0]
    text = re.sub(r"\n{3,}", "\n\n", text)
    b = TextBuilder(ctx.max_chars)
    b.add_raw("全文", text, {})
    t, segs = b.build()
    if not t.strip():
        raise ExtractError("textutil 未能提取到文本")
    return Document(src.name, kind, t, segs, size=src.size, meta={"via": "textutil"})


@register("doc", ".doc", ".rtf", ".wordml", ".webarchive", ".rtfd")
def extract_legacy_doc(src: Source, ctx: ExtractContext) -> Document:
    return _textutil_document(src, ctx, "doc")


@register("xls", ".xls")
def extract_xls(src: Source, ctx: ExtractContext) -> Document:
    try:
        import xlrd  # type: ignore
    except ImportError:
        raise ExtractError("需要 xlrd 解析 .xls：pip install xlrd")
    with src.materialized() as path:
        book = xlrd.open_workbook(path, on_demand=True)
        b = TextBuilder(ctx.max_chars)
        try:
            for sh in book.sheets():
                for r in range(sh.nrows):
                    vals = []
                    for c in range(sh.ncols):
                        v = sh.cell_value(r, c)
                        if v not in ("", None):
                            vals.append(f"{xlrd.formula.colname(c)}{r + 1}={v}")
                    if vals:
                        b.add(f"工作表 {sh.name} · 第 {r + 1} 行", " | ".join(vals),
                              {"sheet": sh.name, "row": r + 1})
        finally:
            with contextlib.suppress(Exception):
                book.release_resources()
    t, segs = b.build()
    return Document(src.name, "xls", t, segs, size=src.size, meta={"sheets": book.sheet_names()})


@register("ppt", ".ppt")
def extract_ppt(src: Source, ctx: ExtractContext) -> Document:
    # 老式 ppt：textutil 不支持，退回字符串扫描
    doc = _strings_document(src, ctx, kind="ppt")
    if not doc.text.strip():
        raise ExtractError("无法解析 .ppt（建议先另存为 .pptx）")
    doc.warning = "老式 .ppt 仅做原始字符串扫描，结构定位不可用"
    return doc


# ---------------------------------------------------------------------------
# 数据库：SQLite
# ---------------------------------------------------------------------------

_SQLITE_MAGIC = b"SQLite format 3\x00"


def _sqlite_open(path: str) -> sqlite3.Connection:
    uri = "file:" + path.replace("?", "%3f").replace("#", "%23")
    try:
        con = sqlite3.connect(uri + "?mode=ro", uri=True)
        con.execute("select count(*) from sqlite_master").fetchone()
        return con
    except sqlite3.DatabaseError:
        pass
    con = sqlite3.connect(uri + "?mode=ro&immutable=1", uri=True)
    return con


def _cell_to_text(v: Any) -> str:
    if v is None:
        return "NULL"
    if isinstance(v, bytes):
        try:
            s = v.decode("utf-8")
            if s.isprintable() or "\n" in s:
                return s[:4000]
        except UnicodeDecodeError:
            pass
        return "0x" + v[:64].hex()
    if isinstance(v, (int, float)):
        return str(v)
    return str(v)


@register("sqlite", ".sqlite", ".sqlite3", ".db", ".db3", ".s3db", ".sl3", ".fossil", ".sqlite2")
def extract_sqlite(src: Source, ctx: ExtractContext) -> Document:
    with src.materialized() as path:
        # 先验证魔数，避免把非 sqlite 的 .db 当数据库
        with open(path, "rb") as f:
            head = f.read(16)
        if not head.startswith(_SQLITE_MAGIC):
            raise ExtractError("不是 SQLite 文件（魔数不匹配）")
        try:
            con = _sqlite_open(path)
        except sqlite3.DatabaseError as e:
            raise ExtractError(f"无法打开数据库: {e}")
        b = TextBuilder(ctx.max_chars)
        tables: List[str] = []
        try:
            rows = con.execute(
                "SELECT type, name, sql FROM sqlite_master WHERE type IN ('table','view') ORDER BY name"
            ).fetchall()
            # 结构
            schema_parts = []
            for typ, name, sql in rows:
                if sql:
                    schema_parts.append(f"-- {typ} {name}\n{sql}")
            if schema_parts:
                b.add("数据库结构 (sqlite_master)", "\n\n".join(schema_parts), {"schema": True})
            for typ, name, sql in rows:
                if not name or name.startswith("sqlite_") and name not in ("sqlite_sequence",):
                    if name != "sqlite_sequence":
                        continue
                tables.append(name)
                b.add(f"表结构 {name}", sql or "", {"table": name, "schema": True})
                quoted = name.replace('"', '""')
                cur = None
                has_rowid = False
                try:
                    # 带上 rowid，便于界面精确跳转到该行
                    cur = con.execute(
                        f'SELECT rowid AS __fg_rowid, * FROM "{quoted}" '
                        f'LIMIT {int(ctx.sqlite_max_rows) + 1}')
                    has_rowid = True
                except sqlite3.DatabaseError:
                    try:
                        cur = con.execute(
                            f'SELECT * FROM "{quoted}" LIMIT {int(ctx.sqlite_max_rows) + 1}')
                    except sqlite3.DatabaseError as e:
                        ctx.warn(f"表 {name} 读取失败: {e}")
                        continue
                cols = [d[0] for d in (cur.description or [])]
                if has_rowid:
                    cols = cols[1:]
                n = 0
                for n, rec in enumerate(cur, 1):
                    if n > ctx.sqlite_max_rows:
                        ctx.warn(f"表 {name} 超过 {ctx.sqlite_max_rows} 行，已截断")
                        break
                    rowid = rec[0] if has_rowid else None
                    vals = rec[1:] if has_rowid else rec
                    pieces = []
                    values: Dict[str, str] = {}
                    for col, val in zip(cols, vals):
                        sval = _cell_to_text(val)
                        values[col] = sval[:512]
                        pieces.append(f"{col}={sval}")
                    meta: Dict[str, Any] = {"table": name, "row": n, "columns": values}
                    if rowid is not None:
                        meta["rowid"] = rowid
                    b.add(f"表 {name} · 第 {n} 行", " | ".join(pieces), meta)
                    if b.truncated:
                        break
                if b.truncated:
                    break
        finally:
            with contextlib.suppress(Exception):
                con.close()
        if b.truncated:
            ctx.warn("数据库内容超长已截断")
        t, segs = b.build()
        doc = Document(src.name, "sqlite", t, segs, size=src.size,
                       meta={"tables": tables, "table_count": len(tables)})
        if not t.strip():
            raise ExtractError("数据库为空或无法读取")
        return doc


# ---------------------------------------------------------------------------
# Access (.mdb/.accdb)：需要 mdbtools
# ---------------------------------------------------------------------------

@register("mdb", ".mdb", ".accdb")
def extract_mdb(src: Source, ctx: ExtractContext) -> Document:
    exe = shutil.which("mdb-tables") or shutil.which("mdb-export")
    if not exe:
        doc = _strings_document(src, ctx, kind="mdb")
        doc.warning = "未安装 mdbtools（brew install mdbtools），仅做原始字符串扫描"
        return doc
    with src.materialized() as path:
        out = _run(["mdb-tables", "-1", path], ctx.timeout).decode("utf-8", "replace").splitlines()
        b = TextBuilder(ctx.max_chars)
        for tbl in (t.strip() for t in out):
            if not tbl:
                continue
            try:
                csv_text = _run(["mdb-export", path, tbl], ctx.timeout).decode("utf-8", "replace")
            except ExtractError as e:
                ctx.warn(f"表 {tbl}: {e}")
                continue
            for i, line in enumerate(csv_text.splitlines()):
                if i == 0:
                    b.add(f"表 {tbl} · 表头", line, {"table": tbl, "header": True})
                elif line.strip():
                    b.add(f"表 {tbl} · 第 {i} 行", line, {"table": tbl, "row": i})
        t, segs = b.build()
        return Document(src.name, "mdb", t, segs, size=src.size, meta={"via": "mdbtools"})


# ---------------------------------------------------------------------------
# 邮件
# ---------------------------------------------------------------------------

@register("eml", ".eml", ".msg")
def extract_eml(src: Source, ctx: ExtractContext) -> Document:
    if src.suffix == ".msg":
        return _strings_document(src, ctx, kind="msg")
    with src.materialized() as path, open(path, "rb") as f:
        msg = message_from_binary_file(f, policy=email_policy.default)
    b = TextBuilder(ctx.max_chars)
    headers = []
    for key in ("From", "To", "Cc", "Subject", "Date", "Message-ID", "Reply-To"):
        if msg.get(key):
            headers.append(f"{key}: {msg.get(key)}")
    if headers:
        b.add("邮件头", "\n".join(headers), {"header": True})
    idx = 0
    for part in msg.walk():
        ctype = part.get_content_type()
        if part.is_multipart():
            continue
        idx += 1
        try:
            payload = part.get_content()
        except Exception:
            payload = part.get_payload(decode=True) or b""
            if isinstance(payload, bytes):
                payload = decode_bytes(payload)[0]
        if not isinstance(payload, str):
            payload = str(payload)
        if ctype == "text/html":
            payload = html_to_text(payload)
            label = f"正文 HTML #{idx}"
        elif ctype.startswith("text/"):
            label = f"正文 {ctype} #{idx}"
        else:
            label = f"附件 {part.get_filename() or ctype} #{idx}"
            payload = payload[:2000] if ctype.startswith("image/") else payload
        if payload.strip():
            b.add(label, payload, {"content_type": ctype,
                                   "filename": part.get_filename() or ""})
    t, segs = b.build()
    return Document(src.name, "eml", t, segs, size=src.size, meta={"subject": msg.get("Subject", "")})


# ---------------------------------------------------------------------------
# 压缩包
# ---------------------------------------------------------------------------

@register("zip", ".zip", ".jar", ".war", ".whl", ".egg", ".xpi", ".vsix", ".apk", )
def extract_zip(src: Source, ctx: ExtractContext) -> Document:
    if ctx.archive_depth <= 0:
        return Document(src.name, "zip", "", [], size=src.size,
                        warning="未启用「压缩包内部搜索」，已跳过")
    zf_cm = zipfile.ZipFile(src.path) if src.path is not None else zipfile.ZipFile(io.BytesIO(src.data or b""))
    b = TextBuilder(ctx.max_chars)
    with zf_cm as zf:
        infos = [i for i in zf.infolist() if not i.is_dir()]
        infos.sort(key=lambda i: i.filename)
        for info in infos[: ctx.max_archive_members]:
            member = info.filename
            if member.endswith("/"):
                continue
            sub_ctx = ExtractContext(**{**ctx.__dict__, "archive_depth": ctx.archive_depth - 1,
                                        "warnings": []})
            sub_src = Source(f"{src.name}::{member}", data=b"")
            # 只对可解析类型递归，其余跳过（避免无谓内存开销）
            sub_suffix = os.path.splitext(member)[1].lower()
            if sub_suffix in _SKIP_MEMBER_SUFFIXES or info.file_size > 64 * 1024 * 1024:
                continue
            try:
                data = zf.read(info)
            except Exception as e:
                ctx.warn(f"{member}: {e}")
                continue
            sub_src = Source(f"{src.name}::{member}", data=data)
            try:
                sub_doc = extract_source(sub_src, sub_ctx)
            except Exception as e:
                ctx.warn(f"{member}: {e}")
                continue
            if sub_doc.text.strip():
                b.add(f"压缩包内 {member}", sub_doc.text, {"member": member})
        t, segs = b.build()
        return Document(src.name, "zip", t, segs, size=src.size,
                        meta={"members": len(infos)})


# ---------------------------------------------------------------------------
# 原始字符串扫描（任意二进制：E01/镜像/未知库/损坏文件）
# ---------------------------------------------------------------------------

# 与 engine.SKIP_MEMBER_SUFFIXES 保持一致（此模块不反向依赖 engine）
_SKIP_MEMBER_SUFFIXES = {
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".ico", ".tif", ".tiff", ".heic",
    ".mp3", ".mp4", ".avi", ".mkv", ".mov", ".wav", ".flac", ".ogg", ".webm", ".m4a", ".aac",
    ".ttf", ".otf", ".woff", ".woff2", ".eot",
    ".zip", ".7z", ".rar", ".gz", ".bz2", ".xz", ".tar", ".jar", ".apk", ".ipa",
    ".whl", ".egg", ".vsix", ".xpi", ".so", ".dylib", ".dll", ".exe", ".class",
    ".dex", ".o", ".a", ".obj", ".lib", ".pyc",
    ".keystore", ".jks", ".p12", ".pfx", ".der",
}

_ASCII_RE = re.compile(rb"[\x20-\x7e\t]{4,}")
_UTF16_RE = re.compile(r"(?:[\x20-\x7e\u00a0-\uffff]\x00){4,}")


def ascii_strings(data: bytes, minlen: int = 4) -> List[Tuple[int, str]]:
    out: List[Tuple[int, str]] = []
    for m in _ASCII_RE.finditer(data):
        s = m.group().decode("ascii", "replace")
        if len(s) >= minlen:
            out.append((m.start(), s))
    return out


def utf16_strings(data: bytes) -> List[Tuple[int, str]]:
    """UTF-16LE 中文/英文串（按字节偏移*2 定位）。"""
    out: List[Tuple[int, str]] = []
    try:
        text = data.decode("utf-16-le", errors="replace")
    except Exception:
        return out
    for m in _UTF16_RE.finditer(text):
        start = m.start() * 2
        s = m.group().replace("\x00", "")
        if len(s.strip()) >= 3:
            out.append((start, s))
    return out


def _strings_document(src: Source, ctx: ExtractContext, kind: str = "binary") -> Document:
    cap = 64 * 1024 * 1024
    data = src.read(max(0, min(src.size, cap)))
    if len(data) < src.size:
        ctx.warn(f"文件较大，仅扫描前 {cap // 1048576}MB")
    strings = ascii_strings(data)
    if len(strings) < 2000:
        strings += utf16_strings(data)
    strings.sort(key=lambda x: x[0])
    b = TextBuilder(ctx.max_chars)
    if not strings:
        b.add_raw("原始数据", "", {})
        t, segs = b.build()
        return Document(src.name, kind, t, segs, size=src.size)
    chunk: List[str] = []
    chunk_start = strings[0][0]
    for off, s in strings:
        chunk.append(f"[0x{off:08x}] {s}")
        if len(chunk) >= 200:
            b.add(f"原始字符串 · 0x{chunk_start:08x}", "\n".join(chunk),
                  {"offset": chunk_start, "raw": True})
            chunk = []
            chunk_start = off
        if b.truncated:
            break
    if chunk:
        b.add(f"原始字符串 · 0x{chunk_start:08x}", "\n".join(chunk),
              {"offset": chunk_start, "raw": True})
    t, segs = b.build()
    return Document(src.name, kind, t, segs, size=src.size,
                    meta={"strings": len(strings)})


# ---------------------------------------------------------------------------
# 嗅探与总入口
# ---------------------------------------------------------------------------

_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _printable_ratio(data: bytes) -> float:
    if not data:
        return 0.0
    sample = data[:8192]
    good = sum(1 for b in sample if b in b"\t\n\r" or 0x20 <= b <= 0x7E or b >= 0x80)
    return good / len(sample)


def sniff(src: Source) -> Optional[Extractor]:
    head = src.read(8192)
    if not head:
        return None
    if head.startswith(_SQLITE_MAGIC):
        return extract_sqlite
    if head.startswith(b"%PDF"):
        return extract_pdf
    if head.startswith(b"PK\x03\x04"):
        # zip 家族：看内部结构决定
        try:
            buf = io.BytesIO(src.data) if src.data is not None else src.path
            with zipfile.ZipFile(buf) as zf:
                names = zf.namelist()
            if any(n.startswith("word/") for n in names):
                return extract_docx
            if any(n.startswith("xl/") for n in names):
                return extract_xlsx
            if any(n.startswith("ppt/") for n in names):
                return extract_pptx
            if "mimetype" in names or "content.xml" in names:
                return extract_odf
            if "META-INF/container.xml" in names:
                return extract_epub
            return extract_zip
        except zipfile.BadZipFile:
            return None
    if head.startswith(b"{\\rtf"):
        return extract_legacy_doc
    if head.startswith(_OLE_MAGIC):
        return extract_legacy_doc   # .doc/.xls/.ppt/.mdb 由具体失败后兜底
    if _printable_ratio(head) > 0.85 and b"\x00" not in head[:512]:
        return extract_text
    return None


def extract_source(src: Source, ctx: Optional[ExtractContext] = None) -> Document:
    """解析单个 Source。永不抛异常。"""
    ctx = ctx or ExtractContext()
    t0 = time.perf_counter()
    fn = _REGISTRY.get(src.suffix)
    kind = _SUFFIX_KIND.get(src.suffix, "binary")
    doc: Optional[Document] = None
    err: Optional[str] = None
    try:
        if fn is None:
            sniffed = sniff(src)
            if sniffed is not None:
                fn = sniffed
                kind = _KIND_BY_FUNC.get(sniffed, "text")
            elif ctx.raw_strings:
                fn = _strings_document
                kind = "binary"
        if fn is not None:
            doc = fn(src, ctx)
            if fn is _strings_document:
                kind = _SUFFIX_KIND.get(src.suffix, "binary")
    except (ExtractError, zipfile.BadZipFile, ET.ParseError) as e:
        err = str(e)
    except sqlite3.DatabaseError as e:
        err = f"数据库错误: {e}"
    except MemoryError:
        err = "内存不足（文件过大）"
    except subprocess.TimeoutExpired:
        err = "解析超时"
    except Exception as e:  # noqa: BLE001 - 批量扫描保底
        err = f"{type(e).__name__}: {e}"

    if (doc is None and fn is not None and fn is not extract_text
            and src.suffix not in _REGISTRY and src.size <= 64 * 1024 * 1024):
        # 未知后缀：即使 sniff 没判定为文本，也试一次解码，用文本质量分过滤乱码
        try:
            cand = extract_text(src, ctx)
            if cand.text.strip() and _text_score(cand.text[:8000]) > 0.62:
                doc = cand
                err = None
        except Exception:
            pass
    if doc is None and _FALLBACK_CONVERTER is not None and src.size <= 32 * 1024 * 1024:
        # 兜底零：markitdown 万能格式转换（覆盖 WPS/旧版 Office/图片元数据等）
        try:
            fb = _FALLBACK_CONVERTER(src, ctx)
            if fb is not None and fb.text.strip():
                doc = fb
                if err:
                    fb.warning = (fb.warning + "；" if fb.warning else "") + f"已用 markitdown 兜底（原错误: {err}）"
                err = None
        except Exception:
            pass
    if doc is None:
        # 兜底一：当文本处理
        if err and src.suffix not in _REGISTRY:
            try:
                doc = extract_text(src, ctx)
                err = None
            except Exception:
                pass
        # 兜底二：原始字符串扫描
        if doc is None and (ctx.raw_strings or (err and src.suffix and src.suffix not in (".png", ".jpg"))):
            try:
                doc = _strings_document(src, ctx, kind=kind)
                doc.warning = (doc.warning + "；" if doc.warning else "") + f"按原始字符串扫描（{err}）"
                err = None
            except Exception:
                pass
    if doc is None:
        doc = Document(src.name, kind, "", [], size=src.size, error=err or "无法解析")
    doc.path = src.name if src.path is None else src.path
    if src.path is None:
        doc.display_name = src.name
    doc.elapsed = time.perf_counter() - t0
    if ctx.warnings and not doc.warning:
        doc.warning = "；".join(ctx.warnings[:3])
    return doc


_KIND_BY_FUNC: Dict[Extractor, str] = {
    extract_text: "text", extract_docx: "docx", extract_xlsx: "xlsx",
    extract_pptx: "pptx", extract_pdf: "pdf", extract_sqlite: "sqlite",
    extract_zip: "zip", extract_odf: "odf", extract_epub: "epub",
    extract_legacy_doc: "doc", extract_xls: "xls", extract_mdb: "mdb",
    extract_eml: "eml", extract_csv: "csv", _strings_document: "binary",
    extract_ppt: "ppt",
}


def extract_path(path: str, ctx: Optional[ExtractContext] = None) -> Document:
    return extract_source(Source(path), ctx)
