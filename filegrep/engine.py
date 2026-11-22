"""文件枚举 + 正则搜索引擎（多进程并行 + 解析缓存）。

速度设计
--------
1. ``backend="process"``：进程池并行解析，绕开 GIL；每个 worker 只回传轻量命中
   结构，不回传全文，因此进程间通信开销很小。
2. 解析缓存（:mod:`filegrep.cache`）：按 路径+mtime+大小 做键，二次搜索同一目录
   时直接命中缓存，通常快一个数量级。
3. ``os.scandir`` 递归枚举（复用 DirEntry 的 stat 缓存）。
4. 可配置并发度；结果流式回传，GUI 边搜边显示，可随时取消。
"""

from __future__ import annotations

import fnmatch
import multiprocessing as mp
import stat as _statmod
import zipfile
import os
import pickle
import re
import threading
import time
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Iterable, Iterator, List, Optional, Pattern, Sequence, Tuple

from . import extract as ex
from .cache import DocCache, get_cache
from .extract import ExtractContext, Source, extract_source
from .model import Document, Hit

__all__ = ["SearchConfig", "SearchEngine", "FileResult", "SearchReport",
           "build_pattern", "iter_files", "DEFAULT_EXCLUDES", "KIND_GROUPS",
           "survey", "load_document", "regex_flags"]

DEFAULT_EXCLUDES = [
    "**/.git/**", "**/.svn/**", "**/.hg/**", "**/node_modules/**", "**/__pycache__/**",
    "**/.venv/**", "**/venv/**", "**/.idea/**", "**/.mypy_cache/**", "**/.DS_Store",
]

# 只给「解析昂贵」的格式做缓存；纯文本重解析比读缓存更快
CACHE_SUFFIXES = {
    ".docx", ".docm", ".dotx", ".doc", ".rtf", ".xlsx", ".xlsm", ".xltx", ".xls",
    ".pptx", ".pptm", ".ppsx", ".ppt", ".pdf", ".odt", ".ods", ".odp", ".epub",
    ".db", ".sqlite", ".sqlite3", ".db3", ".fossil", ".mdb", ".accdb",
    ".zip", ".7z", ".rar", ".jar", ".apk", ".eml", ".msg", ".epub",
}
CACHE_BIG_TEXT = 8 * 1024 * 1024        # 超过 8MB 的文本也缓存

KIND_GROUPS: Dict[str, Tuple[str, ...]] = {
    "全部": (),
    "文档/Office": (".docx", ".docm", ".doc", ".rtf", ".xlsx", ".xlsm", ".xls",
                    ".pptx", ".pptm", ".ppt", ".odt", ".ods", ".odp", ".pdf", ".epub"),
    "文本": (".txt", ".log", ".md", ".csv", ".tsv", ".json", ".jsonl", ".xml", ".ini",
             ".yaml", ".yml", ".sql", ".html", ".htm", ".rtf"),
    "数据库": (".db", ".sqlite", ".sqlite3", ".db3", ".mdb", ".accdb"),
    "压缩包": (".zip", ".7z", ".rar", ".jar", ".apk", ".tar", ".gz"),
    "邮件": (".eml", ".msg", ".mbox"),
    "代码": (".py", ".js", ".ts", ".java", ".c", ".cpp", ".h", ".go", ".php",
             ".sh", ".ps1", ".bat", ".sql", ".rb", ".rs"),
}


_REGEX_SIGNALS = re.compile(
    r"\\[dDwWsSbBAZzG0-9.]"      # \d \w \. 等转义
    r"|\(\?"                     # (?: (?= (?<!
    r"|\[[^\]]*\]"              # [abc]
    r"|\{\d+(?:,\d*)?\}"       # {8} {1,3}
    r"|\.\*|\.\+|\.\?"       # .* .+ .?
    r"|^\^"                       # 开头 ^
    r"|(?<!\\)\$"               # 结尾 $
    r"|(?<!\\)\|"               # a|b
)


def looks_like_regex(pattern: str) -> bool:
    """粗略判断输入是否包含正则语法；没有就当普通文字处理。"""
    if not pattern:
        return False
    return bool(_REGEX_SIGNALS.search(pattern))


def pattern_has_upper(pat: str) -> bool:
    """pattern 里是否出现字面大写字母（跳过 \\A 这类转义）。"""
    i = 0
    n = len(pat)
    while i < n:
        c = pat[i]
        if c == "\\":
            i += 2
            continue
        if "A" <= c <= "Z":
            return True
        i += 1
    return False


@dataclass
class SearchConfig:
    pattern: str = ""
    roots: List[str] = field(default_factory=list)
    ignore_case: bool = True
    case_mode: str = "smart"          # smart | ignore | sensitive
    whole_word: bool = False
    dotall: bool = True
    multiline: bool = True
    literal: bool = False             # 兼容旧字段；match_mode 优先
    match_mode: str = "smart"         # smart | text | regex
    whitelist_exts: Optional[Sequence[str]] = None
    exclude_globs: List[str] = field(default_factory=lambda: list(DEFAULT_EXCLUDES))
    include_hidden: bool = True
    follow_symlinks: bool = False
    max_file_size: int = 512 * 1024 * 1024
    max_hits_per_file: int = 300
    max_total_hits: int = 50000
    context: int = 60
    deep_binary: bool = False
    archives: bool = False
    backend: str = "auto"            # auto | process | thread
    workers: int = 0                 # 0 = 自动
    use_cache: bool = True
    cache_dir: Optional[str] = None
    keep_docs: bool = False
    extract: ExtractContext = field(default_factory=ExtractContext)

    def effective_workers(self) -> int:
        if self.workers and self.workers > 0:
            return self.workers
        cpu = os.cpu_count() or 4
        return max(1, min(12, max(2, cpu // 2 if cpu > 4 else cpu)))

    def pattern_flags(self) -> int:
        flags = 0
        mode = self.case_mode or ("ignore" if self.ignore_case else "sensitive")
        if mode == "ignore":
            flags |= re.IGNORECASE
        elif mode == "smart":
            # 像 ripgrep：pattern 里有大写字母就区分大小写
            if not pattern_has_upper(self.pattern_source()):
                flags |= re.IGNORECASE
        # mode == "sensitive" 不加
        if self.multiline:
            flags |= re.MULTILINE
        if self.dotall:
            flags |= re.DOTALL
        return flags

    def is_literal(self) -> bool:
        """本次是否按纯文字匹配。"""
        mode = self.match_mode or ("text" if self.literal else "regex")
        if mode == "text":
            return True
        if mode == "regex":
            return self.literal          # 显式字面量优先
        return not looks_like_regex(self.pattern)

    def pattern_source(self) -> str:
        pat = self.pattern
        if self.is_literal():
            pat = re.escape(pat)
        if self.whole_word:
            pat = rf"(?<![0-9A-Za-z_])(?:{pat})(?![0-9A-Za-z_])"
        return pat


def regex_flags(cfg: "SearchConfig") -> int:
    return cfg.pattern_flags()


def build_pattern(cfg: SearchConfig) -> Pattern[str]:
    return re.compile(cfg.pattern_source(), cfg.pattern_flags())


# ---------------------------------------------------------------------------
# 枚举
# ---------------------------------------------------------------------------

def _is_excluded(rel: str, name: str, patterns: Sequence[str]) -> bool:
    for raw in patterns:
        p = (raw or "").strip().replace("\\", "/")
        if not p or p.startswith("#"):
            continue
        if p.startswith("**/") and p.endswith("/**"):
            mid = p[3:-3]
            if mid and mid in rel.split("/"):
                return True
        elif "/" in p:
            if fnmatch.fnmatch(rel, p.lstrip("./")) or fnmatch.fnmatch(rel, p):
                return True
        else:
            if fnmatch.fnmatch(name, p):
                return True
    return False


# 只展开真正的 .zip（apk/whl 这类应用包成员太多，会拖慢普通目录扫描）
ZIP_LIKE = {".zip"}
MAX_ZIP_MEMBERS = 5000


# 明确不含可读文本的成员类型；其它一律尝试解析（日志常见 aplog.01 / .1 这类后缀）
SKIP_MEMBER_SUFFIXES = {
    # 图片
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".ico", ".tif", ".tiff", ".heic",
    # 音视频
    ".mp3", ".mp4", ".avi", ".mkv", ".mov", ".wav", ".flac", ".ogg", ".webm", ".m4a", ".aac",
    # 字体
    ".ttf", ".otf", ".woff", ".woff2", ".eot",
    # 压缩包 / 可执行 / 字节码
    ".zip", ".7z", ".rar", ".gz", ".bz2", ".xz", ".tar", ".jar", ".apk", ".ipa",
    ".whl", ".egg", ".vsix", ".xpi", ".so", ".dylib", ".dll", ".exe", ".class",
    ".dex", ".o", ".a", ".obj", ".lib", ".pyc",
    # 证书 / 密钥容器
    ".keystore", ".jks", ".p12", ".pfx", ".der",
}


def _member_worth_parsing(name: str) -> bool:
    """压缩包成员是否值得解析：只跳过明确不可读的类型。"""
    base = name.rsplit("/", 1)[-1]
    suffix = os.path.splitext(base)[1].lower()
    return suffix not in SKIP_MEMBER_SUFFIXES


def _expand_zip(path: str, rel: str, st, cfg: SearchConfig) -> List[Tuple[str, str, int, int, Optional[str]]]:
    """把一个 zip 展开成成员级任务；失败或无可解析成员时返回空列表。"""
    out: List[Tuple[str, str, int, int, Optional[str]]] = []
    try:
        with zipfile.ZipFile(path) as zf:
            for info in zf.infolist()[:MAX_ZIP_MEMBERS]:
                if info.is_dir() or info.file_size <= 0:
                    continue
                if info.file_size > cfg.max_file_size:
                    continue
                if not _member_worth_parsing(info.filename):
                    continue
                try:
                    mt = int(time.mktime(info.date_time + (0, 0, 0)) * 1e9)
                except Exception:
                    mt = st.st_mtime_ns
                out.append((path, rel, info.file_size, mt, info.filename))
    except (zipfile.BadZipFile, OSError):
        return []
    return out


def iter_files(cfg: SearchConfig) -> Iterator[Tuple[str, str, int, int, Optional[str]]]:
    """产出 (绝对路径, 相对显示路径, 大小, mtime_ns, 压缩包成员名或 None)。

    ``archives=True`` 时，zip 家族会被展开成「成员级任务」，这样每个成员都能
    独立并行解析、独立定位，也不会把几百 MB 日志拼成一个巨型字符串。
    """
    seen: set = set()
    whitelist = None
    if cfg.whitelist_exts:
        whitelist = {e.lower() if e.startswith(".") else "." + e.lower() for e in cfg.whitelist_exts}

    def keep(path: str, rel: str, st) -> Optional[Tuple[str, str, int, int, Optional[str]]]:
        if not _statmod.S_ISREG(st.st_mode):
            return None
        if st.st_size > cfg.max_file_size:
            return None
        key = (st.st_dev, st.st_ino)
        if key in seen:
            return None
        seen.add(key)
        if not cfg.include_hidden and any(p.startswith(".") for p in rel.split(os.sep)):
            return None
        if _is_excluded(rel.replace(os.sep, "/"), os.path.basename(path), cfg.exclude_globs):
            return None
        if whitelist is not None:
            if os.path.splitext(path)[1].lower() not in whitelist:
                return None
        return path, rel, st.st_size, st.st_mtime_ns, None

    for root in cfg.roots:
        root = os.path.abspath(os.path.expanduser(root))
        if os.path.isfile(root):
            try:
                st = os.stat(root)
            except OSError:
                continue
            r = keep(root, os.path.basename(root), st)
            if r:
                members = _expand_zip(root, os.path.basename(root), st, cfg) \
                    if cfg.archives and os.path.splitext(root)[1].lower() in ZIP_LIKE else []
                if members:
                    yield from members
                else:
                    yield r
            continue
        if not os.path.isdir(root):
            continue
        stack = [root]
        while stack:
            d = stack.pop()
            try:
                entries = list(os.scandir(d))
            except OSError:
                continue
            entries.sort(key=lambda e: e.name)
            subdirs = []
            for ent in entries:
                try:
                    is_dir = ent.is_dir(follow_symlinks=cfg.follow_symlinks)
                except OSError:
                    continue
                rel = os.path.relpath(ent.path, root)
                if is_dir:
                    if _is_excluded(rel.replace(os.sep, "/"), ent.name, cfg.exclude_globs):
                        continue
                    subdirs.append(ent.path)
                    continue
                try:
                    st = ent.stat(follow_symlinks=cfg.follow_symlinks)
                except OSError:
                    continue
                r = keep(ent.path, rel, st)
                if not r:
                    continue
                member_suffix = os.path.splitext(ent.path)[1].lower()
                if cfg.archives and member_suffix in ZIP_LIKE:
                    members = _expand_zip(ent.path, rel, st, cfg)
                    if members:
                        yield from members
                        continue
                yield r
            stack.extend(reversed(subdirs))


def survey(roots: Sequence[str], excludes: Optional[Sequence[str]] = None) -> Dict[str, int]:
    cfg = SearchConfig(roots=list(roots), exclude_globs=list(excludes or DEFAULT_EXCLUDES))
    stats: Dict[str, int] = {}
    for _path, rel, _s, _m, _mem in iter_files(cfg):
        suffix = os.path.splitext(rel)[1].lower() or "(无扩展名)"
        stats[suffix] = stats.get(suffix, 0) + 1
    return stats


# ---------------------------------------------------------------------------
# 匹配
# ---------------------------------------------------------------------------

def _line_snippet(text: str, start: int, end: int, line_start: int, line_end: int,
                  limit: int = 240) -> Tuple[str, int]:
    if line_end - line_start <= limit * 2:
        return text[line_start:line_end], start - line_start
    lo = max(line_start, start - limit // 2)
    hi = min(line_end, end + limit // 2)
    snippet = text[lo:hi]
    prefix = "…" if lo > line_start else ""
    return prefix + snippet + ("…" if hi < line_end else ""), start - lo + len(prefix)


def search_document(doc: Document, pattern: Pattern[str], cfg: SearchConfig,
                    *, display_name: Optional[str] = None) -> List[Hit]:
    text = doc.text
    if not text:
        return []
    hits: List[Hit] = []
    ctx_len = max(0, cfg.context)
    for m in pattern.finditer(text):
        start, end = m.span()
        if start == end:
            end = min(len(text), start + 1)
        line_no = doc.line_of(start)
        line_start, line_end = doc.line_bounds(line_no)
        seg = doc.segment_at(start)
        snippet, col = _line_snippet(text, start, end, line_start, line_end)
        matched = m.group(0)
        if len(matched) > 500:
            matched = matched[:500] + "…"
        hits.append(Hit(
            path=doc.path,
            display_name=display_name or doc.display_name,
            kind=doc.kind,
            start=start, end=end,
            matched=matched,
            line_no=line_no,
            line_text=snippet,
            col=col,
            label=seg.label if seg else "全文",
            meta=dict(seg.meta) if seg else {},
            before=text[max(0, start - ctx_len):start].replace("\n", " "),
            after=text[end:end + ctx_len].replace("\n", " "),
            text_len=len(text),
        ))
        if len(hits) >= cfg.max_hits_per_file:
            break
    return hits


# ---------------------------------------------------------------------------
# 结果
# ---------------------------------------------------------------------------

@dataclass
class FileResult:
    path: str
    display_name: str
    kind: str
    hits: List[Hit] = field(default_factory=list)
    error: Optional[str] = None
    warning: Optional[str] = None
    elapsed: float = 0.0
    chars: int = 0
    truncated: bool = False
    cached: bool = False
    doc: Optional[Document] = None


@dataclass
class SearchReport:
    results: List[FileResult] = field(default_factory=list)
    files_scanned: int = 0
    files_matched: int = 0
    total_hits: int = 0
    total_files: int = 0
    errors: int = 0
    warnings: int = 0
    cached_files: int = 0
    encrypted: int = 0
    elapsed: float = 0.0
    cancelled: bool = False
    backend: str = ""

    def describe(self) -> str:
        b = f"[{self.backend}] " if self.backend else ""
        extra = f"，{self.encrypted} 个加密无法读取" if self.encrypted else ""
        return (f"{b}扫描 {self.files_scanned}/{self.total_files} 个文件，"
                f"命中 {self.files_matched} 文件 / {self.total_hits} 处，"
                f"缓存 {self.cached_files}{extra}，用时 {self.elapsed:.1f}s")


# ---------------------------------------------------------------------------
# 单文件扫描（线程 / 进程共用）
# ---------------------------------------------------------------------------

def scan_one(path: str, rel: str, st_size: int, st_mtime_ns: int,
             pattern: Pattern[str], cfg: SearchConfig, cache: Optional[DocCache],
             ctx: Optional[ExtractContext] = None,
             member: Optional[str] = None) -> FileResult:
    t0 = time.perf_counter()

    # ---- 压缩包成员：单独解压、单独解析、单独定位 ----
    if member is not None:
        display = member                      # 成员路径比 zip 全名更有辨识度
        doc_path = f"{path}::{member}"
        base_ctx = ctx or cfg.extract
        dctx = ExtractContext(**{**base_ctx.__dict__, "warnings": []})
        dctx.raw_strings = cfg.deep_binary
        dctx.archive_depth = 0
        try:
            with zipfile.ZipFile(path) as zf:
                try:
                    data = zf.read(member)
                except RuntimeError as re_err:
                    if "encrypt" not in str(re_err).lower():
                        raise
                    data = None
                    for pwd in dctx.zip_passwords:
                        try:
                            data = zf.read(member, pwd=pwd.encode("utf-8")
                                           if isinstance(pwd, str) else pwd)
                            break
                        except Exception:                   # noqa: BLE001
                            continue
                    if data is None:
                        return FileResult(path=doc_path, display_name=display, kind="unknown",
                                          error="压缩包成员已加密，需要口令")
        except Exception as e:                              # noqa: BLE001
            return FileResult(path=doc_path, display_name=display, kind="unknown",
                              error=f"读取失败: {e}")
        doc = extract_source(Source(doc_path, data=data), dctx)
        hits = search_document(doc, pattern, cfg, display_name=display)
        for h in hits:
            h.path = doc_path
            h.meta = {**h.meta, "member": member, "archive": path,
                      "archive_name": os.path.basename(path)}
        return FileResult(path=doc_path, display_name=display, kind=doc.kind, hits=hits,
                          error=doc.error, warning=doc.warning,
                          elapsed=time.perf_counter() - t0,
                          chars=len(doc.text), cached=False,
                          truncated=len(hits) >= cfg.max_hits_per_file,
                          doc=doc if cfg.keep_docs else None)

    doc: Optional[Document] = None
    cached = False
    variant = f"a{int(cfg.archives)}d{int(cfg.deep_binary)}m{cfg.extract.max_chars}"
    suffix = os.path.splitext(path)[1].lower()
    cacheable = (suffix in CACHE_SUFFIXES or st_size >= CACHE_BIG_TEXT)
    if cache is not None and cacheable:
        doc = cache.get(path, st_mtime_ns, st_size, variant)
        cached = doc is not None
    if doc is None:
        base_ctx = ctx or cfg.extract
        dctx = ExtractContext(**{**base_ctx.__dict__, "warnings": []})
        dctx.raw_strings = cfg.deep_binary
        dctx.archive_depth = 1 if cfg.archives else 0
        doc = extract_source(Source(path), dctx)
        if cache is not None and cacheable and doc.text and not doc.error:
            cache.put(doc, st_mtime_ns, variant)
    if not doc.text:
        return FileResult(path=path, display_name=rel, kind=doc.kind,
                          error=doc.error, warning=doc.warning,
                          elapsed=doc.elapsed or (time.perf_counter() - t0),
                          chars=0, cached=cached)
    hits = search_document(doc, pattern, cfg, display_name=rel)
    return FileResult(path=path, display_name=rel, kind=doc.kind, hits=hits,
                      error=doc.error, warning=doc.warning,
                      elapsed=doc.elapsed or (time.perf_counter() - t0),
                      chars=len(doc.text), cached=cached,
                      truncated=len(hits) >= cfg.max_hits_per_file,
                      doc=doc if cfg.keep_docs else None)


# ---------------------------------------------------------------------------
# 进程池 worker
# ---------------------------------------------------------------------------

_W: Dict[str, Any] = {}


def _worker_init(cfg_blob: bytes, cancel_flag: Any) -> None:
    cfg: SearchConfig = pickle.loads(cfg_blob)
    _W["cfg"] = cfg
    _W["pattern"] = build_pattern(cfg)
    _W["cancel"] = cancel_flag
    _W["cache"] = get_cache(cfg.use_cache, cfg.cache_dir)
    _W["ctx"] = cfg.extract


def _worker_task(item: Tuple[str, str, int, int, Optional[str]]) -> Optional[FileResult]:
    cancel = _W.get("cancel")
    if cancel is not None and cancel.is_set():
        return None
    cfg: SearchConfig = _W["cfg"]
    path, rel, size, mtime_ns, member = item
    try:
        return scan_one(path, rel, size, mtime_ns, _W["pattern"], cfg, _W["cache"],
                        _W["ctx"], member)
    except Exception as e:                                   # noqa: BLE001
        return FileResult(path=path, display_name=rel, kind="unknown", error=str(e))


# ---------------------------------------------------------------------------
# 引擎
# ---------------------------------------------------------------------------

class SearchEngine:
    def __init__(self, cfg: SearchConfig):
        self.cfg = cfg
        self.pattern = build_pattern(cfg)
        self._cancel = threading.Event()
        self._mp_cancel = None
        self.total_hits = 0
        self.cache = get_cache(cfg.use_cache, cfg.cache_dir)

    def cancel(self) -> None:
        self._cancel.set()
        if self._mp_cancel is not None:
            try:
                self._mp_cancel.set()
            except Exception:
                pass

    @property
    def cancelled(self) -> bool:
        return self._cancel.is_set()

    def _pick_backend(self, n_files: int) -> str:
        cfg = self.cfg
        if cfg.backend in ("process", "thread"):
            return cfg.backend
        if n_files < 64 or (os.cpu_count() or 1) <= 2:
            return "thread"
        if not _mp_available():
            return "thread"
        return "process"

    def run(self, on_event: Optional[Callable[[str, Any], None]] = None) -> SearchReport:
        cfg = self.cfg
        t0 = time.perf_counter()
        report = SearchReport()

        def emit(kind: str, payload: Any = None) -> None:
            if on_event:
                try:
                    on_event(kind, payload)
                except Exception:
                    pass

        files = list(iter_files(cfg))
        report.total_files = len(files)
        backend = self._pick_backend(len(files))
        report.backend = backend
        emit("start", {"total": len(files), "backend": backend})
        if not files:
            report.elapsed = time.perf_counter() - t0
            emit("done", report)
            return report

        def accept(fr: Optional[FileResult]) -> None:
            if fr is None:
                return
            report.files_scanned += 1
            if fr.cached:
                report.cached_files += 1
            if fr.error:
                report.errors += 1
                if "加密" in fr.error:
                    report.encrypted += 1
                    self.encrypted_files = getattr(self, "encrypted_files", 0) + 1
                    emit("encrypted", fr)
                    return
            if fr.warning:
                report.warnings += 1
            if fr.hits:
                report.files_matched += 1
                self.total_hits += len(fr.hits)
                report.results.append(fr)
                emit("file", fr)
            elif fr.error or fr.warning:
                emit("skip", fr)
            if report.files_scanned % 8 == 0 or report.files_scanned == len(files):
                emit("progress", {"done": report.files_scanned, "total": len(files),
                                  "hits": self.total_hits, "cached": report.cached_files})

        if backend == "process":
            self._run_process(files, accept, emit)
        else:
            self._run_thread(files, accept, emit)

        report.total_hits = self.total_hits
        report.encrypted = getattr(self, "encrypted_files", 0)
        report.cancelled = self._cancel.is_set()
        report.elapsed = time.perf_counter() - t0
        if cfg.use_cache and self.cache is not None:
            try:
                self.cache.trim()
            except Exception:
                pass
        emit("done", report)
        return report

    # -- 线程后端 -------------------------------------------------------
    def _run_thread(self, files, accept, emit) -> None:
        cfg = self.cfg
        workers = min(cfg.effective_workers(), max(1, len(files)))
        ctx = cfg.extract
        with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="fg") as pool:
            futures = {pool.submit(scan_one, path, rel, size, mtime,
                                   self.pattern, cfg, self.cache, ctx, member): (path, rel)
                       for path, rel, size, mtime, member in files}
            for fut in as_completed(futures):
                if self._cancel.is_set():
                    break
                try:
                    fr = fut.result()
                except Exception as e:                        # noqa: BLE001
                    path, rel = futures[fut]
                    fr = FileResult(path=path, display_name=rel, kind="unknown", error=str(e))
                accept(fr)
                if self.total_hits >= cfg.max_total_hits:
                    self._cancel.set()
                    break

    # -- 进程后端 -------------------------------------------------------
    def _run_process(self, files, accept, emit) -> None:
        cfg = self.cfg
        workers = min(cfg.effective_workers(), max(1, len(files) // 12 or 1), mp.cpu_count())
        try:
            mctx = mp.get_context("spawn")
            self._mp_cancel = mctx.Event()
        except Exception:
            self._run_thread(files, accept, emit)
            return
        try:
            cfg_blob = pickle.dumps(cfg, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception:
            self._run_thread(files, accept, emit)
            return
        emit("log", f"启动 {workers} 个进程并行解析")
        try:
            with ProcessPoolExecutor(max_workers=workers, mp_context=mctx,
                                     initializer=_worker_init,
                                     initargs=(cfg_blob, self._mp_cancel)) as pool:
                futures = {pool.submit(_worker_task, item): item for item in files}
                try:
                    for fut in as_completed(futures, timeout=None):
                        if self._cancel.is_set():
                            break
                        try:
                            fr = fut.result()
                        except Exception as e:                # noqa: BLE001
                            item = futures[fut]
                            fr = FileResult(path=item[0], display_name=item[1],
                                            kind="unknown", error=f"worker: {e}")
                        accept(fr)
                        if self.total_hits >= cfg.max_total_hits:
                            self._cancel.set()
                            break
                finally:
                    if self._cancel.is_set():
                        for f in futures:
                            f.cancel()
                        pool.shutdown(wait=False, cancel_futures=True)
        except Exception as e:                                # noqa: BLE001
            emit("log", f"进程池异常（{e}），改用线程重新扫描")
            if not self._cancel.is_set():
                self.total_hits = 0
                self._run_thread(files, accept, emit)


def _mp_available() -> bool:
    try:
        mp.get_context("spawn")
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 按需加载（GUI 预览，优先命中缓存）
# ---------------------------------------------------------------------------

def load_document(path: str, *, cfg: Optional[SearchConfig] = None,
                  use_cache: bool = True) -> Document:
    if "::" in path:
        archive, member = path.split("::", 1)
        ctx = ExtractContext(raw_strings=cfg.deep_binary if cfg else False, archive_depth=0)
        try:
            with zipfile.ZipFile(archive) as zf:
                data = zf.read(member)
        except Exception as e:                              # noqa: BLE001
            return Document(path, "binary", "", [], error=f"无法读取压缩包成员: {e}")
        return extract_source(Source(path, data=data), ctx)

    st = None
    try:
        st = os.stat(path)
    except OSError:
        pass
    cache = get_cache(use_cache, cfg.cache_dir if cfg else None) if use_cache else None
    variant = f"a{int(bool(cfg and cfg.archives))}d{int(bool(cfg and cfg.deep_binary))}m{cfg.extract.max_chars if cfg else 0}"
    try:
        suffix = os.path.splitext(path)[1].lower()
    except Exception:
        suffix = ""
    cacheable = st is not None and (suffix in CACHE_SUFFIXES or st.st_size >= CACHE_BIG_TEXT)
    if cache is not None and st is not None and cacheable:
        doc = cache.get(os.path.abspath(path), st.st_mtime_ns, st.st_size, variant)
        if doc is not None:
            return doc
    ctx = ExtractContext(raw_strings=cfg.deep_binary if cfg else False,
                         archive_depth=1 if (cfg and cfg.archives) else 0,
                         max_chars=cfg.extract.max_chars if cfg else ExtractContext().max_chars)
    doc = extract_source(Source(os.path.abspath(path)), ctx)
    if cache is not None and st is not None and cacheable and doc.text and not doc.error:
        cache.put(doc, st.st_mtime_ns, variant)
    return doc


def load_source(name: str, *, cfg: Optional[SearchConfig] = None) -> Document:
    """按 "zip::member" 形式加载压缩包内成员（无缓存）。"""
    ctx = ExtractContext(raw_strings=cfg.deep_binary if cfg else False,
                         archive_depth=1)
    return extract_source(Source(name), ctx)
