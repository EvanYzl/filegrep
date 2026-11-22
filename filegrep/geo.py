"""从任意文本中提取经纬度坐标。

支持的写法
----------
* 十进制度：``116.397, 39.909`` / ``39.909 116.397`` / ``(116.397,39.909)`` / ``[116.397,39.909]``
* 带方向：``39.909°N, 116.397°E`` / ``N39.909 E116.397``
* 度分秒：``39°54'33"N 116°23'29"E`` / ``39°54′33″N``
* 度分：``39°54.5'N 116°23.5'E``
* 带标签：``lat: 39.909, lng: 116.397`` / ``latitude=39.909 longitude=116.397`` / ``纬度 39.909 经度 116.397``

判定策略（按置信度从高到低）
--------------------------
1. ``high``   —— 有 lat/lon 标签、半球字母（N/S/E/W）或度分秒符号
2. ``medium`` —— 同一行内两个裸小数，一个落在纬度范围、一个落在经度范围
3. ``low``    —— 仅当显式放宽时才产出（当前保留给将来的模糊匹配）

默认要求经纬度出现在**同一行**（``same_line_only=True``）；关闭后允许相邻行
（如 ``latitude: 39.909`` 换行 ``longitude: 116.397``）配对。
"""

from __future__ import annotations

import csv
import io
import math
import re
from dataclasses import dataclass, field
from typing import Iterable, Iterator, List, Optional, Sequence, Tuple

__all__ = ["Coord", "extract_coordinates", "to_csv", "CONF_ORDER"]

CONF_ORDER = {"high": 0, "medium": 1, "low": 2}
CONF_LABEL = {"high": "高", "medium": "中", "low": "低"}

# ---------------------------------------------------------------- 正则

_HEMI = r"[NSEWnsew]"

# 度分秒：39°54'33"N / 39度54分33秒N / 39°54.5′
_DMS_RE = re.compile(
    r"(?P<deg>\d{1,3})\s*[°度]\s*"
    r"(?P<min>\d{1,2}(?:\.\d+)?)\s*['′分]?\s*"
    r"(?:(?P<sec>\d{1,2}(?:\.\d+)?)\s*[\"″秒])?"
    r"\s*(?P<hemi>[NSEWnsew])?"
)

# 带标签：lat: 39.9 / latitude=39.9 / 纬度 39.9
_LAT_LABEL_RE = re.compile(
    r"(?:latitude|\blat\b|纬度)\s*[:=：]?\s*(?P<val>[-+]?\d{1,3}(?:\.\d+)?)"
    r"\s*(?P<hemi>[NSEWnsew])?", re.IGNORECASE)
_LON_LABEL_RE = re.compile(
    r"(?:longitude|\blong\b|\blon\b|\blng\b|经度)\s*[:=：]?\s*(?P<val>[-+]?\d{1,3}(?:\.\d+)?)"
    r"\s*(?P<hemi>[NSEWnsew])?", re.IGNORECASE)

# 裸十进制（必须带小数，避免整数误报；排除时间 06.071 与 IP 192.168.1.88）
_DEC_RE = re.compile(
    # 前缀方向字母必须紧贴数字、且前面不是字母（避免 position 的 n、POI 的 I）
    r"(?:(?<![A-Za-z])(?P<h1>[NSEWnsew])\s*)?"
    r"(?<![\d.:\-])(?P<val>[-+]?\d{1,3}\.\d{1,10})"
    r"(?![\d.:])"
    # 后缀方向字母同样要有边界，且后面不能接字母/数字/点（避免吃掉下一个数的前缀 E）
    r"(?:\s*(?P<h2>[NSEWnsew])(?![A-Za-z0-9.]))?"
)

_LAT_RANGE = (-90.0, 90.0)
_LON_RANGE = (-180.0, 180.0)


# ---------------------------------------------------------------- 数据结构

@dataclass
class Coord:
    lat: float
    lon: float
    line_no: int
    start: int          # 在整篇文本中的偏移（含）
    end: int
    raw: str            # 原始片段
    fmt: str            # decimal / dms / dm / labeled
    conf: str           # high / medium / low

    def as_row(self) -> List[str]:
        return [f"{self.lat:.6f}", f"{self.lon:.6f}", self.fmt,
                CONF_LABEL.get(self.conf, self.conf), str(self.line_no),
                _gcj_note(self.lat, self.lon), self.raw.strip()]

    def __repr__(self) -> str:  # pragma: no cover
        return f"<Coord {self.lat:.6f},{self.lon:.6f} L{self.line_no} {self.fmt}/{self.conf}>"


@dataclass
class _Token:
    value: float
    kind: str        # lat | lon | unknown
    start: int
    end: int
    fmt: str
    conf: str
    raw: str
    line_no: int = 0


# ---------------------------------------------------------------- 工具

def _in_lat(v: float) -> bool:
    return _LAT_RANGE[0] <= v <= _LAT_RANGE[1]


def _in_lon(v: float) -> bool:
    return _LON_RANGE[0] <= v <= _LON_RANGE[1]


def _dms_to_dec(deg: str, minute: Optional[str], sec: Optional[str], hemi: Optional[str]) -> Optional[float]:
    try:
        d = float(deg)
        m = float(minute) if minute else 0.0
        s = float(sec) if sec else 0.0
    except (TypeError, ValueError):
        return None
    if m >= 60 or s >= 60:
        return None
    val = d + m / 60.0 + s / 3600.0
    if hemi and hemi.upper() in ("S", "W"):
        val = -val
    return val


def _kind_from_hemi(hemi: Optional[str]) -> str:
    if not hemi:
        return "unknown"
    return "lat" if hemi.upper() in ("N", "S") else "lon"


def _overlaps(a: Tuple[int, int], spans: Sequence[Tuple[int, int]]) -> bool:
    for s, e in spans:
        if a[0] < e and s < a[1]:
            return True
    return False


def _tokens_in_line(line: str, line_no: int, base: int) -> List[_Token]:
    tokens: List[_Token] = []
    used: List[Tuple[int, int]] = []

    # 1) 度分秒（含半球字母）
    for m in _DMS_RE.finditer(line):
        hemi = m.group("hemi")
        val = _dms_to_dec(m.group("deg"), m.group("min"), m.group("sec"), hemi)
        if val is None:
            continue
        kind = _kind_from_hemi(hemi)
        if kind == "unknown":
            kind = "lat" if _in_lat(val) and abs(val) <= 90 else ("lon" if _in_lon(val) else "unknown")
        if kind == "lon" and not _in_lon(val):
            continue
        if kind == "lat" and not _in_lat(val):
            continue
        tok = _Token(val, kind, base + m.start(), base + m.end(), "dms", "high", m.group(0), line_no)
        tokens.append(tok)
        used.append((m.start(), m.end()))

    # 2) 带标签
    for regex, kind in ((_LAT_LABEL_RE, "lat"), (_LON_LABEL_RE, "lon")):
        for m in regex.finditer(line):
            span = (m.start("val"), m.end())
            if _overlaps(span, used):
                continue
            hemi = m.group("hemi")
            if hemi and _kind_from_hemi(hemi) not in (kind, "unknown"):
                kind2 = _kind_from_hemi(hemi)
            else:
                kind2 = kind
            try:
                val = float(m.group("val"))
            except (TypeError, ValueError):
                continue
            if kind2 == "lat" and not _in_lat(val):
                continue
            if kind2 == "lon" and not _in_lon(val):
                continue
            tokens.append(_Token(val, kind2, base + m.start(), base + m.end(),
                                 "labeled", "high", m.group(0), line_no))
            used.append(span)

    # 3) 裸十进制（可带前后半球字母）
    for m in _DEC_RE.finditer(line):
        span = (m.start(), m.end())
        if _overlaps(span, used):
            continue
        try:
            val = float(m.group("val"))
        except (TypeError, ValueError):
            continue
        hemi = m.group("h1") or m.group("h2")
        kind = _kind_from_hemi(hemi)
        conf = "high" if hemi else "medium"
        if kind == "unknown":
            kind = "unknown"
        if not _in_lon(val):
            continue
        tokens.append(_Token(val, kind, base + m.start(), base + m.end(),
                             "decimal", conf, m.group(0), line_no))
        used.append(span)

    tokens.sort(key=lambda t: t.start)
    return tokens


def _make(la: _Token, lo: _Token, conf: str) -> Optional[Coord]:
    lat, lon = la.value, lo.value
    if not _in_lat(lat) or not _in_lon(lon):
        return None
    start = min(la.start, lo.start)
    end = max(la.end, lo.end)
    fmt = la.fmt if la.fmt == lo.fmt else ("labeled" if "labeled" in (la.fmt, lo.fmt) else la.fmt)
    return Coord(lat=lat, lon=lon, line_no=la.line_no, start=start, end=end,
                 raw="", fmt=fmt, conf=conf)


def _pair_line(tokens: List[_Token], line: str, base: int) -> Tuple[List[Coord], set]:
    out: List[Coord] = []
    used: set = set()

    def text_of(a: _Token, b: _Token) -> str:
        s, e = min(a.start, b.start) - base, max(a.end, b.end) - base
        return line[max(0, s):min(len(line), e)]

    # 1) 明确的 lat / lon
    lats = [t for t in tokens if t.kind == "lat"]
    lons = [t for t in tokens if t.kind == "lon"]
    for la in lats:
        if id(la) in used:
            continue
        best = None
        for lo in lons:
            if id(lo) in used:
                continue
            d = abs(lo.start - la.start)
            if best is None or d < best[0]:
                best = (d, lo)
        if best is None:
            continue
        lo = best[1]
        c = _make(la, lo, "high" if (la.conf == "high" or lo.conf == "high") else "medium")
        if c:
            c.raw = text_of(la, lo)
            out.append(c)
            used.add(id(la))
            used.add(id(lo))

    # 2) 剩余裸数字：一个在纬度范围、一个在经度范围就配对（同行，允许中间隔很远）
    rest = [t for t in tokens if id(t) not in used and t.kind == "unknown"]
    lat_like = [t for t in rest if _in_lat(t.value)]
    lon_like = [t for t in rest if not _in_lat(t.value) and _in_lon(t.value)]
    if lat_like and lon_like:
        n = min(len(lat_like), len(lon_like))
        for i in range(n):
            la, lo = lat_like[i], lon_like[i]
            c = _make(la, lo, "medium")
            if c:
                c.raw = text_of(la, lo)
                out.append(c)
                used.add(id(la))
                used.add(id(lo))

    out.sort(key=lambda c: c.start)
    return out, used


# ---------------------------------------------------------------- 主入口

def _format_gcj_hint(lat: float, lon: float) -> str:
    return ""


def _gcj_note(lat: float, lon: float) -> str:
    """给出一句粗略的地理位置提示（仅用于人工扫读，不保证行政区准确）。"""
    ns = "北纬" if lat >= 0 else "南纬"
    ew = "东经" if lon >= 0 else "西经"
    return f"{ns}{abs(lat):.4f} {ew}{abs(lon):.4f}"


def extract_coordinates(text: str, *, same_line_only: bool = True,
                        min_confidence: str = "medium",
                        max_gap_lines: int = 5) -> List[Coord]:
    """扫描整段文本，返回坐标点列表（按出现位置排序）。

    Args:
        same_line_only: True 时只有同一行内的经+纬才算一个点；
            False 时允许相邻 ``max_gap_lines`` 行内配对（如 lat 与 lon 分行写）。
        min_confidence: high / medium / low，低于该置信度的结果被丢弃。
    """
    if not text:
        return []
    threshold = CONF_ORDER.get(min_confidence, 1)
    results: List[Coord] = []
    pending_lat: Optional[_Token] = None      # 跨行模式下的待配对纬度

    pos = 0
    line_no = 1
    n = len(text)
    while pos <= n:
        nl = text.find("\n", pos)
        end = n if nl < 0 else nl
        line = text[pos:end]
        if line.strip():
            tokens = _tokens_in_line(line, line_no, pos)
            # 先做行内配对：同行写法都按常规处理
            line_coords, used = _pair_line(tokens, line, pos)
            results.extend(line_coords)
            if not same_line_only:
                # 跨行：只剩没被行内配对消费的显式 lat/lon 才参与
                leftover = [t for t in tokens if id(t) not in used]
                lons = [t for t in leftover if t.kind == "lon"]
                if pending_lat is not None and lons:
                    lon_tok = lons[0]
                    if line_no - pending_lat.line_no <= max_gap_lines:
                        c = _make(pending_lat, lon_tok, "high")
                        if c:
                            c.raw = f"{pending_lat.raw} … {lon_tok.raw}"
                            results.append(c)
                            pending_lat = None
                lats = [t for t in leftover if t.kind == "lat"]
                if lats:
                    pending_lat = lats[-1]
                elif pending_lat is not None and line_no - pending_lat.line_no > max_gap_lines:
                    pending_lat = None
        if nl < 0:
            break
        pos = nl + 1
        line_no += 1

    results = [c for c in results if CONF_ORDER.get(c.conf, 9) <= threshold]
    results.sort(key=lambda c: c.start)
    return results


def to_csv(coords: Sequence[Coord]) -> str:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["纬度", "经度", "格式", "置信度", "行号", "粗略位置", "原始片段"])
    for c in coords:
        w.writerow(c.as_row())
    return buf.getvalue()
