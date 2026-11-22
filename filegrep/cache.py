"""解析结果缓存：让第二次搜索快到几乎瞬间返回。

* key = 绝对路径 + mtime_ns + 文件大小 + 解析器版本，文件一改就自动失效
* SQLite + zlib 压缩存储；WAL 模式允许进程池里的多个 worker 并发写
* 总容量上限 + LRU 淘汰；超大文档（>32MB 文本）不缓存
"""

from __future__ import annotations

import contextlib
import hashlib
import os
import pickle
import sqlite3
import threading
import time
import zlib
from typing import Optional

from .model import Document, Segment

__all__ = ["DocCache", "get_cache", "CACHE_VERSION"]

CACHE_VERSION = "2"        # 解析逻辑变化时 +1，全量失效


def _default_dir() -> str:
    base = os.environ.get("FILEGREP_CACHE_DIR")
    if base:
        return base
    if os.name == "nt":
        root = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    elif os.sys.platform == "darwin":  # type: ignore[attr-defined]
        root = os.path.join(os.path.expanduser("~"), "Library", "Caches")
    else:
        root = os.environ.get("XDG_CACHE_HOME", os.path.join(os.path.expanduser("~"), ".cache"))
    return os.path.join(root, "filegrep")


class DocCache:
    """线程/进程安全的文档缓存。"""

    def __init__(self, directory: Optional[str] = None, *, max_bytes: int = 1_500_000_000,
                 enabled: bool = True, max_doc_chars: int = 32_000_000):
        self.enabled = enabled
        self.max_bytes = max_bytes
        self.max_doc_chars = max_doc_chars
        self.dir = directory or _default_dir()
        self.db_path = os.path.join(self.dir, "docs.db")
        self._lock = threading.Lock()
        self._local = threading.local()
        self._ok = False
        if enabled:
            with contextlib.suppress(Exception):
                os.makedirs(self.dir, exist_ok=True)
                self._init_db()
                self._ok = True

    # -- 连接管理 -------------------------------------------------------
    def _conn(self) -> sqlite3.Connection:
        conn = getattr(self._local, "conn", None)
        if conn is None:
            conn = sqlite3.connect(self.db_path, timeout=15.0, isolation_level=None)
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA synchronous=NORMAL")
            conn.execute("PRAGMA busy_timeout=15000")
            self._local.conn = conn
        return conn

    def _init_db(self) -> None:
        with self._lock:
            conn = sqlite3.connect(self.db_path, timeout=15.0, isolation_level=None)
            try:
                conn.execute("PRAGMA journal_mode=WAL")
                conn.execute("""
                    CREATE TABLE IF NOT EXISTS docs (
                        key   TEXT PRIMARY KEY,
                        path  TEXT NOT NULL,
                        mtime INTEGER NOT NULL,
                        size  INTEGER NOT NULL,
                        kind  TEXT,
                        nchar INTEGER,
                        nseg  INTEGER,
                        used  REAL,
                        blob  BLOB
                    )""")
                conn.execute("CREATE INDEX IF NOT EXISTS idx_docs_path ON docs(path)")
                conn.execute("CREATE INDEX IF NOT EXISTS idx_docs_used ON docs(used)")
            finally:
                conn.close()

    # -- 键 -------------------------------------------------------------
    @staticmethod
    def make_key(path: str, mtime_ns: int, size: int, variant: str = "") -> str:
        raw = f"{CACHE_VERSION}|{variant}|{os.path.abspath(path)}|{mtime_ns}|{size}"
        return hashlib.sha1(raw.encode("utf-8", "surrogateescape")).hexdigest()

    # -- 读写 -----------------------------------------------------------
    def get(self, path: str, mtime_ns: int, size: int, variant: str = "") -> Optional[Document]:
        if not self._ok:
            return None
        key = self.make_key(path, mtime_ns, size, variant)
        try:
            row = self._conn().execute(
                "SELECT kind, blob FROM docs WHERE key=?", (key,)).fetchone()
        except Exception:
            return None
        if not row:
            return None
        try:
            payload = zlib.decompress(row[1])
            kind, text, segs_meta, meta = pickle.loads(payload)
            segs = [Segment(l, s, e, m) for (l, s, e, m) in segs_meta]
            doc = Document(path, kind or row[0] or "text", text, segs, size=size,
                           mtime=mtime_ns / 1e9, meta=meta or {})
            doc.meta["cached"] = True
            return doc
        except Exception:
            return None

    def put(self, doc: Document, mtime_ns: int, variant: str = "") -> None:
        if not self._ok or len(doc.text) > self.max_doc_chars or not doc.text:
            return
        key = self.make_key(doc.path, mtime_ns, doc.size, variant)
        try:
            payload = pickle.dumps(
                (doc.kind, doc.text,
                 [(s.label, s.start, s.end, s.meta) for s in doc.segments],
                 {k: v for k, v in doc.meta.items() if k != "cached"}),
                protocol=pickle.HIGHEST_PROTOCOL)
            blob = zlib.compress(payload, 1)
        except Exception:
            return
        try:
            self._conn().execute(
                "INSERT OR REPLACE INTO docs(key,path,mtime,size,kind,nchar,nseg,used,blob)"
                " VALUES(?,?,?,?,?,?,?,?,?)",
                (key, os.path.abspath(doc.path), mtime_ns, doc.size, doc.kind,
                 len(doc.text), len(doc.segments), time.time(), blob))
        except Exception:
            return

    # -- 维护 -----------------------------------------------------------
    def stats(self) -> dict:
        if not self._ok:
            return {"enabled": False}
        try:
            row = self._conn().execute(
                "SELECT COUNT(*), COALESCE(SUM(LENGTH(blob)),0) FROM docs").fetchone()
            return {"enabled": True, "entries": row[0], "bytes": row[1], "dir": self.dir}
        except Exception:
            return {"enabled": False}

    def trim(self) -> int:
        """超出容量时按最久未用淘汰。返回删除条数。"""
        if not self._ok:
            return 0
        removed = 0
        try:
            conn = self._conn()
            total = conn.execute("SELECT COALESCE(SUM(LENGTH(blob)),0) FROM docs").fetchone()[0]
            if total <= self.max_bytes:
                return 0
            target = int(self.max_bytes * 0.8)
            rows = conn.execute(
                "SELECT key, LENGTH(blob) FROM docs ORDER BY used ASC").fetchall()
            for key, n in rows:
                if total <= target:
                    break
                conn.execute("DELETE FROM docs WHERE key=?", (key,))
                total -= n or 0
                removed += 1
        except Exception:
            pass
        return removed

    def clear(self) -> None:
        if not self._ok:
            return
        with contextlib.suppress(Exception):
            self._conn().execute("DELETE FROM docs")

    def close(self) -> None:
        conn = getattr(self._local, "conn", None)
        if conn is not None:
            with contextlib.suppress(Exception):
                conn.close()
            self._local.conn = None


_GLOBAL: Optional[DocCache] = None
_GLOBAL_LOCK = threading.Lock()


def get_cache(enabled: bool = True, directory: Optional[str] = None) -> DocCache:
    global _GLOBAL
    if _GLOBAL is None:
        with _GLOBAL_LOCK:
            if _GLOBAL is None:
                _GLOBAL = DocCache(directory, enabled=enabled)
    return _GLOBAL
