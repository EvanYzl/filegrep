"""数据结构：文档、结构单元、命中结果。

设计要点
--------
每个文件被抽取成一段连续的 ``text``（全文），同时保留一组 ``Segment``
记录"哪一段文本来自哪里"（第 3 页 / 工作表 Sheet1 第 12 行 / 表 users 第 5 行 /
压缩包内 member.txt ...）。正则匹配直接在全文上跑，再用二分查找把命中位置
映射回 Segment，从而既支持跨段匹配，又能精确定位到页码/单元格/数据库行。
"""

from __future__ import annotations

import bisect
import os
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

__all__ = ["Segment", "Document", "Hit", "TextBuilder"]


@dataclass(slots=True)
class Segment:
    """全文中的一段连续文本，对应文件里的一个可定位单元。"""

    label: str                       # 人类可读位置，如 "第 3 页"、"工作表 Sheet1 · 第 12 行"
    start: int                       # 在全文中的起始偏移（含）
    end: int                         # 在全文中的结束偏移（不含）
    meta: Dict[str, Any] = field(default_factory=dict)   # 结构化定位信息

    def __repr__(self) -> str:  # pragma: no cover - 调试用
        return f"<Segment {self.label!r} [{self.start}:{self.end}]>"


class Document:
    """一个被解析过的文件。"""

    __slots__ = ("path", "display_name", "kind", "text", "segments", "size", "mtime",
                 "error", "warning", "elapsed", "meta", "_line_starts", "_seg_starts")

    def __init__(self, path: str, kind: str, text: str, segments: List[Segment],
                 *, display_name: Optional[str] = None, size: int = 0, mtime: float = 0.0,
                 error: Optional[str] = None, warning: Optional[str] = None,
                 meta: Optional[Dict[str, Any]] = None):
        self.path = path
        self.display_name = display_name or os.path.basename(path)
        self.kind = kind
        self.text = text
        self.segments = segments
        self.size = size
        self.mtime = mtime
        self.error = error
        self.warning = warning
        self.elapsed = 0.0
        self.meta: Dict[str, Any] = meta or {}
        self._line_starts: Optional[List[int]] = None
        self._seg_starts: Optional[List[int]] = None

    # ---- 行/位置换算 -------------------------------------------------
    @property
    def line_starts(self) -> List[int]:
        ls = self._line_starts
        if ls is None:
            ls = [0]
            text = self.text
            idx = text.find("\n")
            while idx != -1:
                ls.append(idx + 1)
                idx = text.find("\n", idx + 1)
            self._line_starts = ls
        return ls

    @property
    def line_count(self) -> int:
        return len(self.line_starts)

    def line_of(self, pos: int) -> int:
        """偏移 -> 1 基行号。"""
        return bisect.bisect_right(self.line_starts, pos)

    def line_bounds(self, lineno: int) -> Tuple[int, int]:
        """1 基行号 -> (行首偏移, 行尾偏移[不含换行])。"""
        starts = self.line_starts
        lineno = max(1, min(lineno, len(starts)))
        start = starts[lineno - 1]
        end = starts[lineno] - 1 if lineno < len(starts) else len(self.text)
        return start, max(start, end)

    def segment_at(self, pos: int) -> Optional[Segment]:
        """偏移 -> 所属 Segment。"""
        if not self.segments:
            return None
        starts = self._seg_starts
        if starts is None:
            starts = [s.start for s in self.segments]
            self._seg_starts = starts
        i = bisect.bisect_right(starts, pos) - 1
        if i < 0:
            return None
        seg = self.segments[i]
        if pos <= seg.end:
            return seg
        # 落在段间分隔符上，归属前一段
        return seg

    def __repr__(self) -> str:  # pragma: no cover
        return f"<Document {self.display_name} kind={self.kind} chars={len(self.text)} segs={len(self.segments)}>"


@dataclass
class Hit:
    """一次正则命中。"""

    path: str                     # 文件绝对路径（压缩包内成员用 "zip::member" 形式）
    display_name: str
    kind: str                     # 文件类型
    start: int                    # 全文偏移
    end: int
    matched: str                  # 命中的原始文本（可能被截断）
    line_no: int                  # 1 基行号
    line_text: str                # 所在行（截断展示）
    col: int                      # 相对行首的列偏移
    label: str                    # 结构位置，如 "工作表 Sheet1 · 第 12 行"
    meta: Dict[str, Any] = field(default_factory=dict)
    before: str = ""              # 命中前文
    after: str = ""               # 命中后文
    text_len: int = 0             # 该文档全文长度（用于前端估算）

    @property
    def sort_key(self) -> Tuple[str, int]:
        return (self.path, self.start)


class TextBuilder:
    """拼接全文并记录 Segment 的辅助类。"""

    __slots__ = ("_parts", "_pos", "segments", "limit", "truncated")

    def __init__(self, limit: int = 0) -> None:
        self._parts: List[str] = []
        self._pos = 0
        self.segments: List[Segment] = []
        self.limit = max(0, int(limit or 0))
        self.truncated = False

    def add(self, label: str, text: Optional[str], meta: Optional[Dict[str, Any]] = None,
            *, strip: bool = True) -> None:
        if not text:
            return
        if strip:
            text = text.strip("\r\n")
            if not text.strip():
                return
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        if self.limit:
            room = self.limit - self._pos
            if room <= 0:
                self.truncated = True
                return
            if len(text) > room:
                text = text[:room]
                self.truncated = True
        if self._pos:
            self._parts.append("\n")
            self._pos += 1
        start = self._pos
        self._parts.append(text)
        self._pos += len(text)
        self.segments.append(Segment(label, start, self._pos, meta or {}))

    def add_raw(self, label: str, text: str, meta: Optional[Dict[str, Any]] = None) -> None:
        self.add(label, text, meta, strip=False)

    @property
    def length(self) -> int:
        return self._pos

    @property
    def over_limit(self) -> bool:
        return self.truncated

    def build(self) -> Tuple[str, List[Segment]]:
        return "".join(self._parts), self.segments

    def __len__(self) -> int:
        return self._pos
