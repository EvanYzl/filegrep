"""filegrep —— 全格式文件内容检索器（Office / PDF / 数据库 / 压缩包 / 文本）。

对外入口::

    from filegrep import SearchConfig, SearchEngine, load_document

命令行::

    python -m filegrep "正则" /path/to/dir

图形界面::

    python -m filegrep --gui
"""

from __future__ import annotations

__version__ = "1.0.0"

from .model import Document, Hit, Segment          # noqa: E402
from . import extract                              # noqa: E402
from . import parsers                              # noqa: E402

parsers.install()                                  # 把成熟库实现挂到注册表上

from .engine import (                              # noqa: E402
    SearchConfig, SearchEngine, FileResult, SearchReport,
    KIND_GROUPS, build_pattern, load_document, survey,
)
from .cache import DocCache, get_cache             # noqa: E402

__all__ = [
    "Document", "Hit", "Segment", "SearchConfig", "SearchEngine", "FileResult",
    "SearchReport", "DocCache", "get_cache", "load_document", "build_pattern",
    "survey", "KIND_GROUPS", "extract", "parsers", "__version__",
]
