"""python -m filegrep [--gui] ..."""

from __future__ import annotations

import multiprocessing
import sys


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    gui = False
    if argv and argv[0] in ("--gui", "-g", "--window"):
        gui = True
        argv = argv[1:]
    elif not argv:
        gui = True
    if gui:
        from .gui_qt import main as gui_main
        return gui_main(argv)
    from .cli import main as cli_main
    return cli_main(argv)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    raise SystemExit(main())
