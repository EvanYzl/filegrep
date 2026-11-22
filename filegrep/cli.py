"""命令行入口：filegrep [-i] [-w] [-F] PATTERN PATH..."""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from typing import List, Optional

from .engine import KIND_GROUPS, SearchConfig, SearchEngine, survey
from .extract import ExtractContext, supported_suffixes


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="filegrep",
        description="全格式文件内容检索：Office / PDF / 数据库 / 压缩包 / 邮件 / 文本，支持正则",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例:
  filegrep "身份证|银行卡" ~/Downloads
  filegrep -F "192.168.1.1" . --type 全部
  filegrep "(?i)admin.*password" /data --ext .txt,.docx,.db --workers 8
  filegrep --gui           # 打开图形界面
  filegrep --list-types    # 查看支持的文件类型
""")
    p.add_argument("pattern", nargs="?", default="", help="正则表达式")
    p.add_argument("paths", nargs="*", default=[], help="文件或目录")
    p.add_argument("-i", "--ignore-case", action="store_true", default=True, help="忽略大小写（默认开）")
    p.add_argument("-s", "--case-sensitive", dest="ignore_case", action="store_false", help="区分大小写")
    p.add_argument("--case", dest="case_mode", default=None,
                   choices=["smart", "ignore", "sensitive"],
                   help="大小写模式（默认 smart：含大写字母就区分）")
    p.add_argument("-w", "--word", action="store_true", help="整词匹配")
    p.add_argument("-F", "--literal", action="store_true", help="按字面量匹配（不当正则）")
    p.add_argument("--no-multiline", dest="multiline", action="store_false", default=True,
                   help="^ $ 不按行匹配")
    p.add_argument("--no-dotall", dest="dotall", action="store_false", default=True,
                   help=". 不匹配换行")
    p.add_argument("--type", dest="type_group", default="全部", help=f"类型分组: {', '.join(KIND_GROUPS)}")
    p.add_argument("--ext", default="", help="扩展名白名单，逗号分隔，如 .txt,.docx,.db")
    p.add_argument("--exclude", action="append", default=[], help="排除 glob，可多次")
    p.add_argument("--max-hits", type=int, default=300, help="单文件最多命中数")
    p.add_argument("--context", type=int, default=60, help="上下文长度")
    p.add_argument("--workers", type=int, default=0, help="并发度（0=自动）")
    p.add_argument("--backend", default="auto", choices=["auto", "process", "thread"], help="并行后端")
    p.add_argument("--no-cache", action="store_true", help="禁用解析缓存")
    p.add_argument("--deep", action="store_true", help="对二进制文件做原始字符串扫描")
    p.add_argument("--archives", dest="archives", action="store_true", default=True,
                   help="搜索压缩包内部（默认开启）")
    p.add_argument("--no-archives", dest="archives", action="store_false",
                   help="不进入压缩包")
    p.add_argument("--zip-pass", default="", help="加密 zip 口令，多个用逗号分隔")
    p.add_argument("--max-size", type=int, default=512, help="最大文件大小(MB)")
    p.add_argument("--json", action="store_true", help="输出 JSON")
    p.add_argument("-c", "--count", action="store_true", help="只输出计数")
    p.add_argument("-l", "--files-only", action="store_true", help="只输出命中的文件名")
    p.add_argument("--stats", action="store_true", help="显示统计信息")
    p.add_argument("--list-types", action="store_true", help="列出支持的文件类型并退出")
    p.add_argument("--no-color", action="store_true", help="禁用颜色")
    return p


_COLORS = {"path": "\033[1;36m", "loc": "\033[33m", "match": "\033[1;31m", "dim": "\033[2m", "reset": "\033[0m"}


def _highlight(text: str, hit, color: bool) -> str:
    if not color:
        return text
    rel = hit.col
    n = len(hit.matched)
    if rel < 0 or rel + n > len(text):
        return text
    return text[:rel] + _COLORS["match"] + text[rel:rel + n] + _COLORS["reset"] + text[rel + n:]


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    if args.list_types:
        print(f"共支持 {len(supported_suffixes())} 种扩展名:\n")
        line: List[str] = []
        for s in supported_suffixes():
            line.append(s)
            if len(line) >= 12:
                print("  " + "  ".join(f"{x:<9}" for x in line))
                line = []
        if line:
            print("  " + "  ".join(f"{x:<9}" for x in line))
        return 0

    if not args.pattern:
        build_parser().print_help()
        return 2
    paths = args.paths or ["."]
    for path in paths:
        if not os.path.exists(path):
            print(f"路径不存在: {path}", file=sys.stderr)
            return 2

    exts = None
    if args.ext:
        exts = [e if e.startswith(".") else "." + e for e in args.ext.split(",") if e.strip()]
    else:
        exts = list(KIND_GROUPS.get(args.type_group, ())) or None

    if args.case_mode:
        case_mode = args.case_mode
    elif not args.ignore_case:
        case_mode = "sensitive"
    else:
        case_mode = "smart"

    cfg = SearchConfig(
        pattern=args.pattern,
        case_mode=case_mode,
        roots=paths,
        ignore_case=args.ignore_case,
        whole_word=args.word,
        literal=args.literal,
        multiline=args.multiline,
        dotall=args.dotall,
        whitelist_exts=exts,
        exclude_globs=__import__("filegrep.engine", fromlist=["DEFAULT_EXCLUDES"]).DEFAULT_EXCLUDES
        + list(args.exclude),
        max_hits_per_file=args.max_hits,
        context=args.context,
        workers=args.workers,
        backend=args.backend,
        use_cache=not args.no_cache,
        deep_binary=args.deep,
        archives=args.archives,
        extract=ExtractContext(
            zip_passwords=[p for p in args.zip_pass.split(",") if p]),
        max_file_size=args.max_size * 1024 * 1024,
    )

    color = not args.no_color and sys.stdout.isatty()
    if args.json:
        collected = []

    def on_event(kind: str, payload) -> None:
        if kind == "encrypted":
            if args.stats:
                print(f"{_COLORS['dim']}加密跳过 {payload.display_name}{_COLORS['reset']}", file=sys.stderr)
            return
        if kind in ("file", "skip"):
            fr = payload
            if args.json:
                if fr.hits:
                    collected.append({
                        "path": fr.path,
                        "display": fr.display_name,
                        "kind": fr.kind,
                        "elapsed": round(fr.elapsed, 3),
                        "error": fr.error,
                        "hits": [{
                            "line": h.line_no, "col": h.col, "label": h.label,
                            "match": h.matched, "text": h.line_text,
                            "span": [h.start, h.end], "meta": h.meta,
                        } for h in fr.hits],
                    })
            elif fr.hits:
                if args.files_only:
                    print(fr.display_name)
                elif not args.count:
                    for h in fr.hits:
                        loc = f"{_COLORS['path']}{fr.display_name}{_COLORS['reset']}:{h.line_no}:{h.col + 1}"
                        lab = f"{_COLORS['dim']}[{h.label}]{_COLORS['reset']}" if h.label else ""
                        print(f"{loc}: {lab} {_highlight(h.line_text.strip(), h, color)}")
            elif fr.error and args.stats:
                print(f"{_COLORS['dim']}跳过 {fr.display_name}: {fr.error}{_COLORS['reset']}", file=sys.stderr)
        elif kind == "log" and args.stats:
            print(f"{_COLORS['dim']}{payload}{_COLORS['reset']}", file=sys.stderr)

    engine = SearchEngine(cfg)
    report = engine.run(on_event)

    if args.json:
        print(json.dumps({
            "files_scanned": report.files_scanned,
            "files_matched": report.files_matched,
            "total_hits": report.total_hits,
            "elapsed": round(report.elapsed, 3),
            "backend": report.backend,
            "results": collected,
        }, ensure_ascii=False, indent=2))
    elif args.count:
        print(report.files_matched)
    if args.stats:
        print(report.describe(), file=sys.stderr)
    return 0 if report.total_hits else 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
