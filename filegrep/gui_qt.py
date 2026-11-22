"""PySide6 图形界面：文件夹全格式检索 + 浏览器 + 数据库查询。

三个面板
--------
* 左：结果树 / 文件浏览器（可浏览目录下任意类型的文件）
* 右：文本预览（行号、命中高亮、大文件滚动增量加载）/ 数据库 SQL 控制台
* 下：命中详情（页码 / 工作表单元格 / 数据库表行）

交互
----
* 双击命中：预览跳到该位置并高亮，全部命中淡黄、当前命中橙红
* 滚动到底部：自动加载下一段（超大文件也能顺畅滑动）
* ⌘Shift+R 在 Finder 中显示，⌘O 用系统程序打开，F3 / ⇧F3 上下一个命中
* 选中 .db/.sqlite：右侧「数据库」页自动可用，写 SQL、点表名、导出结果
"""

from __future__ import annotations

import csv
import json
import os
import re
import sqlite3
import subprocess
import sys
from typing import Any, Dict, List, Optional

from PySide6.QtCore import (QFileInfo, QObject, QRect, QSettings, QSize, Qt, QThread,
                            QTimer, QUrl, Signal, Slot)
from PySide6.QtGui import (QAction, QColor, QFont, QFontDatabase, QFontMetrics, QIcon,
                           QKeySequence, QPainter, QPalette, QShortcut, QTextCharFormat,
                           QTextCursor)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox, QComboBox,
                               QFileDialog, QFileIconProvider, QHBoxLayout, QHeaderView,
                               QLabel, QLineEdit, QListWidget, QListWidgetItem,
                               QMainWindow, QMenu, QMessageBox, QPlainTextEdit,
                               QProgressBar, QPushButton, QSizePolicy, QSpinBox,
                               QSplitter, QStyle, QStyledItemDelegate,
                               QStyleOptionViewItem, QTabWidget, QTableWidget,
                               QTableWidgetItem, QTextBrowser, QTextEdit, QToolButton, QTreeWidget,
                               QTreeWidgetItem, QVBoxLayout, QWidget)

from .engine import (DEFAULT_EXCLUDES, KIND_GROUPS, FileResult, SearchConfig,
                     SearchEngine, SearchReport, load_document, looks_like_regex)
from .extract import ExtractContext, supported_suffixes

APP_NAME = "FileGrep"
APP_TITLE = "FileGrep · 全格式文件检索 / 浏览 / 数据库查询"

HIT_BG = QColor("#fdf3d3")
HIT_FG = QColor("#4a3f18")
CUR_BG = QColor("#ffd28a")
CUR_FG = QColor("#4a3200")

# 清爽风：白面板 + 淡灰分隔 + 低饱和蓝作为唯一强调色
STYLE = """
QMainWindow, QDialog { background: #fbfcfd; }
QStatusBar { background: #fbfcfd; color: #6b7280; }
QStatusBar::item { border: none; }
QMenuBar { background: #fbfcfd; color: #37414b; }
QMenuBar::item { padding: 4px 10px; background: transparent; }
QMenuBar::item:selected { background: #eef3f9; border-radius: 5px; }
QMenu { background: #ffffff; border: 1px solid #e6eaef; padding: 4px; }
QMenu::item { padding: 5px 22px 5px 14px; border-radius: 5px; }
QMenu::item:selected { background: #eef3f9; }
QLineEdit, QComboBox, QSpinBox { background: #ffffff; border: 1px solid #e3e8ee;
    border-radius: 7px; padding: 4px 9px; min-height: 20px; selection-background-color: #cfe3ff;
    selection-color: #1f2328; }
QLineEdit:focus, QComboBox:focus, QSpinBox:focus { border-color: #a9c4e2; }
QLineEdit:disabled, QComboBox:disabled { color: #aab1b9; background: #f7f9fb; }
QPushButton { background: #ffffff; border: 1px solid #e3e8ee; border-radius: 7px;
    padding: 4px 14px; color: #2b3238; min-height: 20px; }
QPushButton:hover { background: #f5f8fb; border-color: #d2dbe4; }
QPushButton:pressed { background: #eef2f7; }
QPushButton:disabled { color: #b3b9c0; background: #f8fafc; border-color: #eceff3; }
QPushButton#primary { background: #5b8fc9; border: 1px solid #5b8fc9; color: #ffffff; }
QPushButton#primary:hover { background: #4f83bd; border-color: #4f83bd; }
QPushButton#primary:pressed { background: #4577ad; }
QPushButton#primary:disabled { background: #c6d6e8; border-color: #c6d6e8; color: #ffffff; }
QTreeWidget, QListWidget, QTableWidget, QTextBrowser, QPlainTextEdit {
    background: #ffffff; border: 1px solid #e6eaef; border-radius: 9px; }
QTreeWidget::item, QListWidget::item { height: 25px; padding-left: 2px; }
QTreeWidget::item:hover, QListWidget::item:hover { background: #f5f8fc; }
QTreeWidget::item:selected, QListWidget::item:selected { background: #e8f0fa; color: #1f2328; }
QHeaderView::section { background: #f8fafc; border: none; border-bottom: 1px solid #e9edf2;
    color: #6b7280; padding: 6px 8px; }
QTableWidget { gridline-color: #eef1f5; selection-background-color: #e8f0fa;
    selection-color: #1f2328; }
QTabWidget::pane { border: 1px solid #e6eaef; border-radius: 9px; background: #ffffff; top: -1px; }
QTabBar::tab { background: transparent; color: #6b7280; padding: 6px 15px; margin-right: 2px; }
QTabBar::tab:selected { color: #3f6ea8; }
QTabBar::tab:hover { color: #37414b; }
QProgressBar { border: none; background: #edf1f6; border-radius: 3px; max-height: 6px; }
QProgressBar::chunk { background: #9dbde0; border-radius: 3px; }
QSplitter::handle { background: transparent; }
QCheckBox { spacing: 6px; color: #37414b; }
QToolButton { background: transparent; border: 1px solid transparent; border-radius: 6px;
    padding: 3px 9px; color: #5f6b76; }
QToolButton:hover { background: #f1f5f9; border-color: #e3e8ee; }
QLabel#hint { color: #6b7280; }
QToolButton#chip { background: #eef3f9; border: 1px solid #dbe4ee; border-radius: 10px;
    padding: 2px 9px; color: #33475b; font-size: 12px; }
QToolButton#chip:hover { background: #e3ecf6; border-color: #c9d8e8; }
QToolButton#chipExclude { background: #fdeeee; border: 1px solid #f2d4d4; border-radius: 10px;
    padding: 2px 9px; color: #a03a3a; font-size: 12px; }
QToolButton#chipExclude:hover { background: #fbdede; border-color: #e9c2c2; }
QLabel#filterInfo { color: #3f6ea8; }
QLabel#title { color: #55606b; }
"""

MONO_CANDIDATES = ("SF Mono", "Menlo", "Consolas", "DejaVu Sans Mono", "Monaco", "Courier New")
UI_CANDIDATES = ("PingFang SC", "Microsoft YaHei", "Noto Sans CJK SC", "Source Han Sans SC")
NONBMP = re.compile(r"[\U00010000-\U0010FFFF]")

DB_SUFFIXES = {".db", ".sqlite", ".sqlite3", ".db3", ".s3db", ".sl3", ".fossil"}


def pick_font(candidates, fallback: str) -> str:
    fams = set(QFontDatabase.families())
    for c in candidates:
        if c in fams:
            return c
    return fallback


def human_size(n: int) -> str:
    if n < 1024:
        return f"{n}B"
    for unit in ("KB", "MB", "GB"):
        n /= 1024.0
        if n < 1024:
            return f"{n:.1f}{unit}"
    return f"{n:.1f}TB"


def safe_text(s: str) -> str:
    """把非 BMP 字符替换成 1 个 BMP 字符，保证 Python 索引 == Qt 位置。"""
    return NONBMP.sub("▯", s) if NONBMP.search(s) else s


# ---------------------------------------------------------------------------
# 预览编辑器
# ---------------------------------------------------------------------------

class LineNumberArea(QWidget):
    def __init__(self, editor: "PreviewEdit"):
        super().__init__(editor)
        self._editor = editor

    def sizeHint(self) -> QSize:
        return QSize(self._editor.line_number_area_width(), 0)

    def paintEvent(self, event):
        self._editor.paint_line_numbers(event)


class PreviewEdit(QPlainTextEdit):
    """只读预览：行号栏、命中高亮、滚动到底自动加载下一段。"""

    needMore = Signal()
    filterRequested = Signal(str)
    excludeRequested = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setUndoRedoEnabled(False)
        self.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        font = QFont(pick_font(MONO_CANDIDATES, "Menlo"))
        font.setStyleHint(QFont.StyleHint.Monospace)
        font.setPointSize(12)
        self.setFont(font)
        self._gutter = LineNumberArea(self)
        self.blockCountChanged.connect(self._update_width)
        self.updateRequest.connect(self._on_update)
        self._update_width()
        self._gutter_bg = QColor("#f8fafc")
        self._gutter_fg = QColor("#b3b9c0")
        self._loading_more = False
        self._line_offset = 0
        self.verticalScrollBar().valueChanged.connect(self._check_bottom)
        self._at_end = False

    # -- 行号 -----------------------------------------------------------
    def line_number_area_width(self) -> int:
        digits = max(3, len(str(max(1, self.blockCount()))))
        return 16 + self.fontMetrics().horizontalAdvance("9") * digits

    def _update_width(self) -> None:
        self.setViewportMargins(self.line_number_area_width(), 0, 0, 0)

    def _on_update(self, rect: QRect, dy: int) -> None:
        if dy:
            self._gutter.scroll(0, dy)
        else:
            self._gutter.update(0, rect.y(), self._gutter.width(), rect.height())
        if rect.contains(self.viewport().rect()):
            self._update_width()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        cr = self.contentsRect()
        self._gutter.setGeometry(QRect(cr.left(), cr.top(),
                                       self.line_number_area_width(), cr.height()))

    def paint_line_numbers(self, event) -> None:
        painter = QPainter(self._gutter)
        painter.fillRect(event.rect(), self._gutter_bg)
        block = self.firstVisibleBlock()
        num = block.blockNumber()
        offset = self.contentOffset()
        top = self.blockBoundingGeometry(block).translated(offset).top()
        bottom = top + self.blockBoundingRect(block).height()
        painter.setPen(self._gutter_fg)
        font = self.font()
        font.setPointSize(max(9, font.pointSize() - 2))
        painter.setFont(font)
        width = self._gutter.width() - 8
        height = self.fontMetrics().height()
        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                painter.drawText(0, int(top), width, height,
                                 int(Qt.AlignmentFlag.AlignRight),
                                 str(num + 1 + self._line_offset))
            block = block.next()
            top = bottom
            bottom = top + self.blockBoundingRect(block).height()
            num += 1

    # -- 增量加载 -------------------------------------------------------
    def _check_bottom(self, value: int) -> None:
        if self._at_end:
            return
        sb = self.verticalScrollBar()
        if sb.maximum() - value <= 3 and sb.maximum() > 0:
            self._at_end = True
            QTimer.singleShot(0, self._emit_more)

    def _emit_more(self) -> None:
        self._at_end = False
        self.needMore.emit()

    def contextMenuEvent(self, event) -> None:
        menu = self.createStandardContextMenu()
        sel = self.textCursor().selectedText().replace("\u2029", " ").strip()
        if 0 < len(sel) <= 200:
            menu.addSeparator()
            short = sel[:24] + ("…" if len(sel) > 24 else "")
            act = menu.addAction(f'用「{short}」继续筛选（包含）')
            act.triggered.connect(lambda: self.filterRequested.emit(sel))
            act2 = menu.addAction(f'排除含「{short}」的行')
            act2.triggered.connect(lambda: self.excludeRequested.emit(sel))
        menu.exec(event.globalPos())

    def set_line_offset(self, offset: int) -> None:
        self._line_offset = max(0, int(offset))
        self._gutter.update()

    def append_chunk(self, text: str) -> None:
        cur = QTextCursor(self.document())
        cur.movePosition(QTextCursor.MoveOperation.End)
        cur.insertText(text)


# ---------------------------------------------------------------------------
# 内容片段着色
# ---------------------------------------------------------------------------

class SnippetDelegate(QStyledItemDelegate):
    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index) -> None:
        hit = index.data(Qt.ItemDataRole.UserRole + 1)
        if hit is None or not getattr(hit, "line_text", ""):
            super().paint(painter, option, index)
            return
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        text = opt.text
        opt.text = ""
        style = opt.widget.style() if opt.widget else QApplication.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, opt, painter, opt.widget)
        painter.save()
        selected = bool(option.state & QStyle.StateFlag.State_Selected)
        rect = option.rect.adjusted(3, 0, -3, 0)
        fm = painter.fontMetrics()
        base_color = (option.palette.color(QPalette.ColorRole.HighlightedText) if selected
                      else option.palette.color(QPalette.ColorRole.Text))
        col = max(0, min(getattr(hit, "col", 0), len(text)))
        n = max(1, min(len(getattr(hit, "matched", "")), len(text) - col))
        parts = (text[:col], text[col:col + n], text[col + n:])
        x = rect.left()
        y = rect.top() + (rect.height() + fm.ascent() - fm.descent()) // 2
        painter.setPen(base_color)
        for i, seg in enumerate(parts):
            w = fm.horizontalAdvance(seg)
            if x > rect.right():
                break
            if i == 1 and seg:
                painter.fillRect(QRect(x - 1, rect.top() + 1, w + 2, rect.height() - 2),
                                 QColor("#eed9a3") if selected else QColor("#f6e8c4"))
                painter.setPen(QColor("#5b4a1f"))
            painter.drawText(x, y, seg)
            x += w
            if i == 0:
                painter.setPen(base_color)
        painter.restore()


# ---------------------------------------------------------------------------
# 后台线程
# ---------------------------------------------------------------------------

class SearchThread(QThread):
    fileFound = Signal(object)
    progress = Signal(int, int, int, int)
    finishedReport = Signal(object)
    logLine = Signal(str)
    encryptedFound = Signal(object)

    def __init__(self, cfg: SearchConfig, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.engine = SearchEngine(cfg)

    def run(self) -> None:
        report = SearchReport()
        try:
            def on_event(kind: str, payload: Any) -> None:
                if kind == "file":
                    self.fileFound.emit(payload)
                elif kind == "progress":
                    self.progress.emit(payload["done"], payload["total"],
                                       payload["hits"], payload.get("cached", 0))
                elif kind == "log":
                    self.logLine.emit(str(payload))
                elif kind == "encrypted":
                    self.encryptedFound.emit(payload)
                elif kind == "start":
                    self.progress.emit(0, payload["total"], 0, 0)
            report = self.engine.run(on_event)
        except Exception as e:                      # noqa: BLE001
            self.logLine.emit(f"搜索出错: {e}")
        finally:
            self.finishedReport.emit(report)

    def stop(self) -> None:
        self.engine.cancel()


class DocLoader(QThread):
    loaded = Signal(object, int)

    def __init__(self, path: str, cfg: Optional[SearchConfig], token: int, parent=None):
        super().__init__(parent)
        self.path = path
        self.cfg = cfg
        self.token = token

    def run(self) -> None:
        try:
            doc = load_document(self.path, cfg=self.cfg)
        except Exception as e:                      # noqa: BLE001
            self.loaded.emit(e, self.token)
            return
        self.loaded.emit(doc, self.token)


class DirScanner(QThread):
    """后台扫描单个目录，避免大目录卡住界面。"""

    done = Signal(str, list, int)     # path, entries[(name,is_dir,size,mtime)], total
    MAX = 4000

    def __init__(self, path: str, parent=None):
        super().__init__(parent)
        self.path = path

    def run(self) -> None:
        entries: List[tuple] = []
        total = 0
        try:
            with os.scandir(self.path) as it:
                for e in it:
                    total += 1
                    if len(entries) >= self.MAX:
                        continue
                    try:
                        is_dir = e.is_dir()
                        st = e.stat()
                        entries.append((e.name, is_dir, st.st_size, st.st_mtime))
                    except OSError:
                        entries.append((e.name, False, 0, 0))
        except OSError:
            pass
        entries.sort(key=lambda x: (not x[1], x[0].lower()))
        self.done.emit(self.path, entries, total)


# ---------------------------------------------------------------------------
# 结果树
# ---------------------------------------------------------------------------

class ResultTree(QTreeWidget):
    hitActivated = Signal(object)        # 单击：文本预览定位
    hitOpened = Signal(object)           # 双击 / 回车：数据库命中直接跳到对应行

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setColumnCount(4)
        self.setHeaderLabels(["文件", "位置", "行", "内容"])
        self.setUniformRowHeights(True)
        self.setAlternatingRowColors(False)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setRootIsDecorated(True)
        self.setItemDelegateForColumn(3, SnippetDelegate(self))
        self.header().setStretchLastSection(True)
        self.setColumnWidth(0, 250)
        self.setColumnWidth(1, 120)
        self.setColumnWidth(2, 46)
        self.itemActivated.connect(self._on_open)
        self.itemClicked.connect(self._on_item)

    def _on_item(self, item: QTreeWidgetItem, column: int) -> None:
        hit = item.data(0, Qt.ItemDataRole.UserRole + 1)
        if hit is not None:
            self.hitActivated.emit(hit)
        elif item.childCount():
            item.setExpanded(not item.isExpanded())

    def _on_open(self, item: QTreeWidgetItem, column: int) -> None:
        hit = item.data(0, Qt.ItemDataRole.UserRole + 1)
        if hit is not None:
            self.hitOpened.emit(hit)
        elif item.childCount():
            item.setExpanded(not item.isExpanded())


# ---------------------------------------------------------------------------
# 文件浏览器
# ---------------------------------------------------------------------------

class FileBrowser(QWidget):
    fileChosen = Signal(str)          # 单击 → 预览
    fileActivated = Signal(str)       # 双击 → 预览并定位开头

    def __init__(self, parent=None):
        super().__init__(parent)
        self._icons: Dict[str, QIcon] = {}
        self._icon_provider: Optional[QFileIconProvider] = None
        self._scanners: Dict[str, DirScanner] = {}
        self._path_items: Dict[str, QTreeWidgetItem] = {}
        self._root = ""
        self._generation = 0

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        top = QHBoxLayout()
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("按文件名过滤…")
        self.filter_edit.setClearButtonEnabled(True)
        self.filter_edit.textChanged.connect(self._apply_filter)
        top.addWidget(self.filter_edit, 1)
        layout.addLayout(top)

        self.tree = QTreeWidget()
        self.tree.setColumnCount(4)
        self.tree.setHeaderLabels(["名称", "大小", "类型", "修改时间"])
        self.tree.setUniformRowHeights(True)
        self.tree.setAlternatingRowColors(False)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.tree.header().setStretchLastSection(True)
        self.tree.setColumnWidth(0, 270)
        self.tree.setColumnWidth(1, 72)
        self.tree.setColumnWidth(2, 66)
        self.tree.setColumnWidth(3, 96)
        self.tree.itemExpanded.connect(self._on_expand)
        self.tree.itemClicked.connect(self._on_click)
        self.tree.itemDoubleClicked.connect(self._on_double)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._menu)
        layout.addWidget(self.tree, 1)

    # -- 图标 -----------------------------------------------------------
    def _icon(self, path: str, is_dir: bool) -> QIcon:
        key = "dir" if is_dir else (os.path.splitext(path)[1].lower() or "file")
        if key in self._icons:
            return self._icons[key]
        if self._icon_provider is None:
            self._icon_provider = QFileIconProvider()
        try:
            icon = self._icon_provider.icon(QFileInfo(path))
        except Exception:
            icon = QIcon()
        self._icons[key] = icon
        return icon

    # -- 根目录 ---------------------------------------------------------
    def set_root(self, path: str) -> None:
        path = os.path.abspath(os.path.expanduser(path or ""))
        if not os.path.isdir(path) or path == self._root:
            return
        self._root = path
        self._generation += 1
        self.tree.clear()
        self._path_items.clear()
        root_item = QTreeWidgetItem([os.path.basename(path) or path, "", "文件夹", ""])
        root_item.setData(0, Qt.ItemDataRole.UserRole, path)
        root_item.setIcon(0, self._icon(path, True))
        self.tree.addTopLevelItem(root_item)
        root_item.setExpanded(True)
        self._scan(root_item, path)

    def _scan(self, item: QTreeWidgetItem, path: str) -> None:
        if item.childCount() and item.child(0).data(0, Qt.ItemDataRole.UserRole) != "__placeholder__":
            return
        item.takeChildren()
        item.addChild(QTreeWidgetItem(["载入中…"]))
        gen = self._generation
        scanner = DirScanner(path, self)
        scanner.done.connect(
            lambda p, entries, total, it=item, g=gen: self._filled(it, p, entries, total, g))
        scanner.finished.connect(scanner.deleteLater)
        self._scanners[path] = scanner
        scanner.start()

    def _filled(self, item: QTreeWidgetItem, path: str, entries: list, total: int,
                gen: int = -1) -> None:
        from datetime import datetime
        if gen != self._generation:
            return
        try:
            item.takeChildren()
        except RuntimeError:          # C++ 对象已销毁（换过根目录）
            return
        self._path_items[path] = item
        for name, is_dir, size, mtime in entries:
            full = os.path.join(path, name)
            try:
                mt = datetime.fromtimestamp(mtime).strftime("%m-%d %H:%M")
            except (OSError, ValueError, OverflowError):
                mt = ""
            kind = "文件夹" if is_dir else (os.path.splitext(name)[1].lower().lstrip(".") or "文件")
            child = QTreeWidgetItem([name, "" if is_dir else human_size(size), kind, mt])
            child.setData(0, Qt.ItemDataRole.UserRole, full)
            child.setIcon(0, self._icon(full, is_dir))
            if is_dir:
                ph = QTreeWidgetItem(["…"])
                ph.setData(0, Qt.ItemDataRole.UserRole, "__placeholder__")
                child.addChild(ph)
            item.addChild(child)
        if total > len(entries):
            more = QTreeWidgetItem([f"… 共 {total} 项，仅显示前 {len(entries)} 项"])
            more.setFlags(Qt.ItemFlag.NoItemFlags)
            item.addChild(more)
        self._apply_filter(self.filter_edit.text())
        if item.childCount() and item is self.tree.topLevelItem(0):
            item.setExpanded(True)

    def _on_expand(self, item: QTreeWidgetItem) -> None:
        try:
            path = item.data(0, Qt.ItemDataRole.UserRole)
        except RuntimeError:
            return
        if path and path != "__placeholder__":
            self._scan(item, path)

    def _on_click(self, item: QTreeWidgetItem, col: int) -> None:
        path = item.data(0, Qt.ItemDataRole.UserRole)
        if path and path != "__placeholder__" and os.path.isfile(path):
            self.fileChosen.emit(path)

    def _on_double(self, item: QTreeWidgetItem, col: int) -> None:
        path = item.data(0, Qt.ItemDataRole.UserRole)
        if path and path != "__placeholder__" and os.path.isfile(path):
            self.fileActivated.emit(path)

    def _apply_filter(self, text: str) -> None:
        text = (text or "").strip().lower()

        def walk(item: QTreeWidgetItem) -> bool:
            any_visible = False
            for i in range(item.childCount()):
                ch = item.child(i)
                if ch.data(0, Qt.ItemDataRole.UserRole) == "__placeholder__":
                    continue
                hit = walk(ch)
                if text and text not in ch.text(0).lower() and not hit:
                    ch.setHidden(True)
                else:
                    ch.setHidden(False)
                    any_visible = True
            return any_visible

        for i in range(self.tree.topLevelItemCount()):
            walk(self.tree.topLevelItem(i))

    def _menu(self, pos) -> None:
        item = self.tree.itemAt(pos)
        if item is None:
            return
        path = item.data(0, Qt.ItemDataRole.UserRole)
        if not path or path == "__placeholder__":
            return
        menu = QMenu(self)
        act_reveal = menu.addAction("在 Finder 中显示")
        act_open = menu.addAction("用默认程序打开")
        menu.addSeparator()
        act_copy = menu.addAction("复制路径")
        chosen = menu.exec(self.tree.viewport().mapToGlobal(pos))
        if chosen is act_reveal:
            if sys.platform == "darwin":
                subprocess.Popen(["open", "-R", path])
            elif os.name == "nt":
                subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(path)])
        elif chosen is act_open:
            self.fileActivated.emit(path)
        elif chosen is act_copy:
            QApplication.clipboard().setText(path)


# ---------------------------------------------------------------------------
# 数据库查询面板
# ---------------------------------------------------------------------------

class DbPanel(QWidget):
    """SQLite 只读查询控制台。"""

    statusMessage = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.conn: Optional[sqlite3.Connection] = None
        self.path: Optional[str] = None
        self.max_rows = 2000

        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(6)

        left = QVBoxLayout()
        left.setSpacing(4)
        self.db_label = QLabel("未打开数据库")
        self.db_label.setObjectName("hint")
        self.db_label.setWordWrap(True)
        left.addWidget(self.db_label)
        self.tables = QListWidget()
        self.tables.setMaximumWidth(230)
        self.tables.itemDoubleClicked.connect(self._on_table_double)
        self.tables.itemClicked.connect(self._on_table_click)
        left.addWidget(self.tables, 1)
        left_box = QWidget()
        left_box.setLayout(left)
        left_box.setMaximumWidth(220)
        left_box.setMinimumWidth(180)
        root.addWidget(left_box)

        right = QVBoxLayout()
        right.setSpacing(4)
        sql_row = QHBoxLayout()
        self.sql_edit = QPlainTextEdit()
        self.sql_edit.setPlaceholderText('SELECT * FROM users WHERE name LIKE "%张%" LIMIT 100')
        f = QFont(pick_font(MONO_CANDIDATES, "Menlo"))
        f.setPointSize(12)
        self.sql_edit.setFont(f)
        self.sql_edit.setFixedHeight(78)
        sql_row.addWidget(self.sql_edit, 1)
        btns = QVBoxLayout()
        self.btn_run = QPushButton("执行")
        self.btn_run.clicked.connect(self.run_sql)
        self.btn_clear = QPushButton("清空")
        self.btn_clear.clicked.connect(lambda: self.result.setRowCount(0))
        self.btn_export = QPushButton("导出 CSV")
        self.btn_export.clicked.connect(self.export_csv)
        self.btn_run.setToolTip("执行 SQL（⌘Enter）")
        btns.addWidget(self.btn_run)
        btns.addWidget(self.btn_clear)
        btns.addWidget(self.btn_export)
        sql_row.addLayout(btns)
        right.addLayout(sql_row)

        self.info = QLabel("")
        self.info.setObjectName("hint")
        right.addWidget(self.info)

        self.result = QTableWidget()
        self.result.setAlternatingRowColors(False)
        self.result.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.result.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.result.setSortingEnabled(False)
        self.result.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        right.addWidget(self.result, 1)
        root.addLayout(right)

        sc = QShortcut(QKeySequence("Meta+Return"), self.sql_edit, activated=self.run_sql)
        sc2 = QShortcut(QKeySequence("Ctrl+Return"), self.sql_edit, activated=self.run_sql)
        self.setEnabled(False)

    # -- 打开 -----------------------------------------------------------
    def open_db(self, path: str) -> bool:
        self.close_db()
        self.path = path
        try:
            uri = "file:" + path.replace("?", "%3f").replace("#", "%23") + "?mode=ro"
            self.conn = sqlite3.connect(uri, uri=True, timeout=5)
            rows = self.conn.execute(
                "SELECT type, name FROM sqlite_master WHERE type IN ('table','view') "
                "AND name NOT LIKE 'sqlite_%' ORDER BY type, name").fetchall()
        except sqlite3.DatabaseError as e:
            self.db_label.setText(f"无法打开：{e}")
            self.conn = None
            self.setEnabled(False)
            return False
        self.setEnabled(True)
        self.tables.clear()
        for typ, name in rows:
            item = QListWidgetItem(name)
            item.setData(Qt.ItemDataRole.UserRole, name)
            item.setToolTip("表" if typ == "table" else "视图")
            if typ == "view":
                item.setForeground(QColor("#6b7280"))
            self.tables.addItem(item)
        self.db_label.setText(f"{os.path.basename(path)}\n{len(rows)} 张表/视图")
        if self.tables.count():
            self.tables.setCurrentRow(0)
            self.preview_table(self.tables.item(0).data(Qt.ItemDataRole.UserRole))
        return True

    def close_db(self) -> None:
        if self.conn is not None:
            try:
                self.conn.close()
            except Exception:
                pass
        self.conn = None
        self.path = None
        self.tables.clear()
        self.result.setRowCount(0)
        self.info.setText("")
        self.db_label.setText("未打开数据库")
        self.setEnabled(False)

    # -- 查询 -----------------------------------------------------------
    def _on_table_click(self, item: QListWidgetItem) -> None:
        name = item.data(Qt.ItemDataRole.UserRole)
        if name:
            self.sql_edit.setPlainText(f'SELECT * FROM "{name}" LIMIT 200;')

    def _on_table_double(self, item: QListWidgetItem) -> None:
        name = item.data(Qt.ItemDataRole.UserRole)
        if name:
            self.sql_edit.setPlainText(f'SELECT * FROM "{name}" LIMIT 200;')
            self.preview_table(name)

    def preview_table(self, name: str) -> None:
        self.sql_edit.setPlainText(f'SELECT * FROM "{name}" LIMIT 200;')
        self.run_sql()

    def _query(self, sql: str, params=()) -> tuple[list, list, bool]:
        cur = self.conn.execute(sql, params)
        cols = [d[0] for d in (cur.description or [])]
        rows = cur.fetchmany(self.max_rows + 1)
        truncated = len(rows) > self.max_rows
        return cols, rows[: self.max_rows], truncated

    def _fill(self, cols: list, rows: list, info: str) -> None:
        self.result.setUpdatesEnabled(False)
        try:
            self.result.clear()
            self.result.setColumnCount(len(cols))
            self.result.setHorizontalHeaderLabels([str(c) for c in cols])
            self.result.setRowCount(len(rows))
            for r, row in enumerate(rows):
                for c, val in enumerate(row):
                    if isinstance(val, bytes):
                        text = "0x" + val[:24].hex() + ("…" if len(val) > 24 else "")
                    else:
                        text = "" if val is None else str(val)
                    item = QTableWidgetItem(text)
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self.result.setItem(r, c, item)
        finally:
            self.result.setUpdatesEnabled(True)
            self.result.resizeColumnsToContents()
        self.info.setText(info)
        self.statusMessage.emit(info)

    def run_sql(self) -> None:
        if self.conn is None:
            return
        sql = self.sql_edit.toPlainText().strip()
        if not sql:
            return
        for stmt in [x.strip() for x in sql.split(";") if x.strip()]:
            if re.match(r"(?i)^\s*(insert|update|delete|drop|alter|create|replace|attach|pragma\s+writable)",
                        stmt):
                self.info.setText("只读模式：已阻止写操作")
                continue
            try:
                cols, rows, truncated = self._query(stmt)
            except sqlite3.DatabaseError as e:
                self.info.setText(f"SQL 错误：{e}")
                self.result.setRowCount(0)
                self.result.setColumnCount(0)
                return
            info = f"返回 {len(rows)} 行 × {len(cols)} 列"
            if truncated:
                info += f"（已截断，最多 {self.max_rows} 行）"
            self._fill(cols, rows, info)
            break

    def locate(self, table: str, rowid=None, row=None) -> None:
        """跳到指定表的目标行：显示上下各 8 行并选中目标行。"""
        if self.conn is None:
            return
        for i in range(self.tables.count()):
            if self.tables.item(i).data(Qt.ItemDataRole.UserRole) == table:
                self.tables.setCurrentRow(i)
                break
        quoted = table.replace('"', '""')
        result = None
        target = 0
        if rowid is not None:
            try:
                lo = max(0, int(rowid) - 8)
                hi = int(rowid) + 8
                cols, rows, _ = self._query(
                    f'SELECT rowid, * FROM "{quoted}" WHERE rowid BETWEEN ? AND ? ORDER BY rowid',
                    (lo, hi))
                result = (cols, rows)
                # 用返回结果里的 rowid 反查下标，避免 lo 与真实 rowid 起点不一致
                target = 0
                for i, r in enumerate(rows):
                    if r and r[0] == int(rowid):
                        target = i
                        break
                self.sql_edit.setPlainText(
                    f'SELECT rowid, * FROM "{quoted}" WHERE rowid BETWEEN {lo} AND {hi};')
            except (sqlite3.DatabaseError, TypeError, ValueError):
                result = None
        if result is None:
            try:
                r0 = int(row or 1)
                offset = max(0, r0 - 1 - 8)
                cols, rows, _ = self._query(
                    f'SELECT * FROM "{quoted}" LIMIT 17 OFFSET ?', (offset,))
                result = (cols, rows)
                target = max(0, r0 - 1 - offset)
                self.sql_edit.setPlainText(
                    f'SELECT * FROM "{quoted}" LIMIT 17 OFFSET {offset};')
            except (sqlite3.DatabaseError, TypeError, ValueError) as e:
                self.info.setText(f"定位失败：{e}")
                return
        cols, rows = result
        if rows and not (0 <= target < len(rows)):
            target = 0
        extra = f"（rowid={rowid}）" if rowid is not None else ""
        self._fill(cols, rows, f"已定位 {table} 第 {row} 行{extra}，上下各 8 行")
        if 0 <= target < len(rows):
            self.result.selectRow(target)
            item = self.result.item(target, 0)
            if item is not None:
                self.result.scrollToItem(
                    item, QAbstractItemView.ScrollHint.PositionAtCenter)

    def show_schema(self) -> None:
        if self.conn is None:
            return
        item = self.tables.currentItem()
        name = item.data(Qt.ItemDataRole.UserRole) if item else None
        if not name:
            self.info.setText("请先选择一张表")
            return
        self.sql_edit.setPlainText(
            f'SELECT cid, name, type, "notnull", dflt_value, pk FROM pragma_table_info({name!r});')
        self.run_sql()

    def export_csv(self) -> None:
        if self.result.rowCount() == 0:
            self.info.setText("没有可导出的结果")
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出查询结果",
                                              os.path.expanduser("~/query_result.csv"),
                                              "CSV (*.csv)")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            cols = [self.result.horizontalHeaderItem(c).text() for c in range(self.result.columnCount())]
            w.writerow(cols)
            for r in range(self.result.rowCount()):
                w.writerow([self.result.item(r, c).text() if self.result.item(r, c) else ""
                            for c in range(self.result.columnCount())])
        self.info.setText(f"已导出到 {path}")


# ---------------------------------------------------------------------------
# 主窗口
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    PREVIEW_CHUNK = 300_000
    PREVIEW_WINDOW = 300_000

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1480, 940)
        self.setAcceptDrops(True)

        self.settings = QSettings("filegrep", "FileGrep")
        self.search_thread: Optional[SearchThread] = None
        self.loader: Optional[DocLoader] = None
        self.loader_token = 0
        self.current_doc = None
        self.current_hit = None
        self.current_path: Optional[str] = None
        self._file_hits: Dict[str, list] = {}
        self._pending: List[FileResult] = []
        self._flush_timer = QTimer(self)
        self._flush_timer.setInterval(80)
        self._flush_timer.timeout.connect(self._flush_pending)
        self._hit_order: List[Any] = []
        self._hit_index = -1
        self._cfg_last: Optional[SearchConfig] = None
        self._full_text = ""
        self._loaded_chars = 0
        self._browse_mode = False
        self._windowed = False
        self._doc_terms: List[tuple] = []
        self._doc_match_lines: List[tuple] = []
        self._doc_match_index = 0
        self._win_len = 0
        self._last_browse: Optional[str] = None

        self.setStyleSheet(STYLE)
        self._build_ui()
        self._restore_settings()
        self._update_cache_label()

    # -- UI -------------------------------------------------------------
    def _build_ui(self) -> None:
        ui_font = QFont(pick_font(UI_CANDIDATES, ""))
        ui_font.setPointSize(13)
        self.setFont(ui_font)

        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(16, 12, 16, 10)
        root.setSpacing(9)

        # 路径
        row1 = QHBoxLayout()
        row1.setSpacing(8)
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("要搜索的文件夹（可拖拽文件夹到窗口）")
        self.path_edit.setClearButtonEnabled(True)
        self.path_edit.returnPressed.connect(self.on_path_changed)
        self.path_edit.editingFinished.connect(self.on_path_changed)
        btn_browse = QPushButton("浏览…")
        btn_browse.clicked.connect(self.browse_folder)
        row1.addWidget(QLabel("目录"))
        row1.addWidget(self.path_edit, 1)
        row1.addWidget(btn_browse)

        # 正则
        row2 = QHBoxLayout()
        row2.setSpacing(8)
        self.pattern_combo = QComboBox()
        self.pattern_combo.setEditable(True)
        self.pattern_combo.setInsertPolicy(QComboBox.InsertPolicy.InsertAtTop)
        self.pattern_combo.lineEdit().setPlaceholderText(
            r"正则表达式，例如  张三|138\d{8}    或    (?i)password\s*=")
        self.pattern_combo.lineEdit().returnPressed.connect(self.start_search)
        self.pattern_combo.setMinimumHeight(32)
        self.btn_search = QPushButton("搜索")
        self.btn_search.setDefault(True)
        self.btn_search.setObjectName("primary")
        self.btn_search.clicked.connect(self.start_search)
        self.btn_search.setMinimumHeight(32)
        self.btn_stop = QPushButton("停止")
        self.btn_stop.setEnabled(False)
        self.btn_stop.setMinimumHeight(32)
        self.btn_stop.clicked.connect(self.stop_search)
        row2.addWidget(QLabel("正则"))
        row2.addWidget(self.pattern_combo, 1)
        row2.addWidget(self.btn_search)
        row2.addWidget(self.btn_stop)

        # 选项
        row3 = QHBoxLayout()
        row3.setSpacing(12)
        self.case_combo = QComboBox()
        self.case_combo.addItem("智能大小写", "smart")
        self.case_combo.addItem("忽略大小写", "ignore")
        self.case_combo.addItem("区分大小写", "sensitive")
        self.case_combo.setToolTip("智能：正则里出现大写字母就区分大小写（推荐）")
        self.cb_word = QCheckBox("整词")
        self.cb_word.setToolTip("只匹配完整单词（给表达式两端加词边界）")
        self.match_combo = QComboBox()
        self.match_combo.addItem("智能匹配", "smart")
        self.match_combo.addItem("纯文字", "text")
        self.match_combo.addItem("正则", "regex")
        self.match_combo.setToolTip(
            "智能匹配：输入像普通文字就直接找文字，像正则才按正则（推荐）\n"
            "纯文字：所有符号原样查找\n"
            "正则：始终按正则表达式")
        self.cb_deep = QCheckBox("二进制字符串")
        self.cb_deep.setToolTip("对无法解析的二进制文件（镜像、未知库）做可打印字符串扫描")
        self.cb_arch = QCheckBox("压缩包内部")
        self.cb_arch.setChecked(True)
        self.cb_arch.setToolTip("递归搜索 zip/7z/rar 内部文件（较慢）")
        self.cb_cache = QCheckBox("解析缓存")
        self.cb_cache.setChecked(True)
        self.type_combo = QComboBox()
        for name in KIND_GROUPS:
            self.type_combo.addItem(name)
        self.backend_combo = QComboBox()
        self.backend_combo.addItem("并行：自动", "auto")
        self.backend_combo.addItem("并行：多进程", "process")
        self.backend_combo.addItem("并行：多线程", "thread")
        self.backend_combo.setToolTip("多进程绕过 GIL，最快")
        self.workers_spin = QSpinBox()
        self.workers_spin.setRange(0, 32)
        self.workers_spin.setValue(0)
        self.workers_spin.setPrefix("并发 ")
        self.hits_spin = QSpinBox()
        self.hits_spin.setRange(0, 1_000_000)
        self.hits_spin.setSingleStep(50)
        self.hits_spin.setValue(0)
        self.hits_spin.setSpecialValueText("单文件命中 不限")
        self.hits_spin.setPrefix("单文件命中 ")
        self.hits_spin.setToolTip("0 = 不限；需要时再设阈值，避免个别文件命中过多")
        self.zip_pwd_edit = QLineEdit()
        self.zip_pwd_edit.setPlaceholderText("zip 口令")
        self.zip_pwd_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.zip_pwd_edit.setMaximumWidth(130)
        self.zip_pwd_edit.setToolTip("加密 zip 的口令，多个用逗号分隔；填好后回车或点「重搜」")
        self.zip_pwd_edit.returnPressed.connect(self.start_search)
        self.btn_retry_enc = QPushButton("重搜加密项")
        self.btn_retry_enc.setVisible(False)
        self.btn_retry_enc.setToolTip("用当前口令重新搜索被跳过的加密内容")
        self.btn_retry_enc.clicked.connect(self.start_search)
        for w in (self.cb_word, self.cb_deep, self.cb_arch, self.cb_cache):
            row3.addWidget(w)
        row3.addWidget(self.match_combo)
        row3.addWidget(self.case_combo)
        row3.addSpacing(8)
        row3.addWidget(QLabel("类型"))
        row3.addWidget(self.type_combo)
        row3.addWidget(self.backend_combo)
        row3.addWidget(self.workers_spin)
        row3.addWidget(self.hits_spin)
        row3.addWidget(self.zip_pwd_edit)
        row3.addWidget(self.btn_retry_enc)
        row3.addStretch(1)
        root.addLayout(row1)
        root.addLayout(row2)
        root.addLayout(row3)

        # 主体
        splitter = QSplitter(Qt.Orientation.Horizontal)

        self.left_tabs = QTabWidget()
        self.left_tabs.setDocumentMode(True)

        # 结果页
        result_page = QWidget()
        rv = QVBoxLayout(result_page)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(4)
        frow = QHBoxLayout()
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("在结果中过滤：空格分隔多个词，-词 表示排除")
        self.filter_edit.setClearButtonEnabled(True)
        self.filter_edit.textChanged.connect(self.apply_filter)
        self.result_filter_info = QLabel("")
        self.result_filter_info.setObjectName("filterInfo")
        self.result_filter_info.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Preferred)
        btn_expand = QToolButton()
        btn_expand.setText("展开")
        btn_expand.clicked.connect(lambda: self.tree.expandAll())
        btn_collapse = QToolButton()
        btn_collapse.setText("折叠")
        btn_collapse.clicked.connect(lambda: self.tree.collapseAll())
        frow.addWidget(self.filter_edit, 1)
        frow.addWidget(self.result_filter_info)
        frow.addWidget(btn_expand)
        frow.addWidget(btn_collapse)
        rv.addLayout(frow)
        self.tree = ResultTree()
        self.tree.hitActivated.connect(self.show_hit)
        self.tree.hitOpened.connect(self.open_hit_in_db)
        rv.addWidget(self.tree, 1)
        self.left_tabs.addTab(result_page, "结果")

        # 文件浏览器页
        self.browser = FileBrowser()
        self.browser.fileChosen.connect(self.browse_file)
        self.browser.fileActivated.connect(self.browse_file)
        self.left_tabs.addTab(self.browser, "文件")

        splitter.addWidget(self.left_tabs)

        # 右侧
        self.right_tabs = QTabWidget()
        self.right_tabs.setDocumentMode(True)

        preview_page = QWidget()
        pv = QVBoxLayout(preview_page)
        pv.setContentsMargins(0, 0, 0, 0)
        pv.setSpacing(4)
        head = QHBoxLayout()
        self.preview_title = QLabel("预览")
        self.preview_title.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.preview_title.setObjectName("title")
        head.addWidget(self.preview_title, 1)
        self.wrap_check = QCheckBox("自动换行")
        self.wrap_check.toggled.connect(
            lambda on: self.editor.setLineWrapMode(
                QPlainTextEdit.LineWrapMode.WidgetWidth if on else QPlainTextEdit.LineWrapMode.NoWrap))
        head.addWidget(self.wrap_check)
        pv.addLayout(head)

        # 文件内查找条
        find_row = QHBoxLayout()
        find_row.setSpacing(6)
        self.filter_mode_combo = QComboBox()
        self.filter_mode_combo.addItem("包含", "include")
        self.filter_mode_combo.addItem("排除", "exclude")
        self.filter_mode_combo.setToolTip("选择新条件的极性；也可在输入开头写 - 或 ! 直接表示排除")
        self.filter_mode_combo.setMaximumWidth(86)
        self.doc_filter_edit = QLineEdit()
        self.doc_filter_edit.setPlaceholderText(
            "在当前文件里继续查找：回车叠加一个条件（同时满足）；以 - 或 ! 开头表示排除")
        self.doc_filter_edit.setClearButtonEnabled(True)
        self.doc_filter_edit.returnPressed.connect(self.add_doc_filter)
        self.btn_filter_prev = QPushButton("上一处")
        self.btn_filter_prev.clicked.connect(lambda: self.goto_doc_match(-1))
        self.btn_filter_next = QPushButton("下一处")
        self.btn_filter_next.clicked.connect(lambda: self.goto_doc_match(1))
        self.btn_filter_clear = QPushButton("清除")
        self.btn_filter_clear.clicked.connect(self.clear_doc_filters)
        self.filter_info = QLabel("")
        self.filter_info.setObjectName("filterInfo")
        self.filter_info.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.filter_info.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Preferred)
        self.doc_filter_edit.setMinimumWidth(140)
        self.doc_filter_edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        for b in (self.btn_filter_prev, self.btn_filter_next, self.btn_filter_clear):
            b.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        find_row.addWidget(self.filter_mode_combo)
        find_row.addWidget(self.doc_filter_edit, 1)
        find_row.addWidget(self.btn_filter_prev)
        find_row.addWidget(self.btn_filter_next)
        find_row.addWidget(self.btn_filter_clear)
        find_row.setStretch(1, 1)
        pv.addLayout(find_row)

        self.filter_chips = QWidget()
        self.filter_chips_layout = QHBoxLayout(self.filter_chips)
        self.filter_chips_layout.setContentsMargins(0, 0, 0, 0)
        self.filter_chips_layout.setSpacing(6)
        self.filter_chips_layout.addStretch(1)
        self.filter_chips_layout.addWidget(self.filter_info)
        pv.addWidget(self.filter_chips)

        self.editor = PreviewEdit()
        self.editor.needMore.connect(self._load_more)
        self.editor.filterRequested.connect(self.add_doc_filter)
        self.editor.excludeRequested.connect(self.add_doc_filter_exclude)
        pv.addWidget(self.editor, 1)

        self.detail = QTextBrowser()
        self.detail.setOpenExternalLinks(True)
        self.detail.setMinimumHeight(76)
        self.detail.setMaximumHeight(230)
        pv.addWidget(self.detail)
        self.right_tabs.addTab(preview_page, "文本预览")

        self.db_panel = DbPanel()
        self.db_panel.statusMessage.connect(lambda s: self.statusBar().showMessage(s, 5000))
        self.right_tabs.addTab(self.db_panel, "数据库")

        splitter.addWidget(self.right_tabs)
        splitter.setSizes([680, 800])
        root.addWidget(splitter, 1)

        # 底部
        bottom = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setMaximumWidth(280)
        self.progress.setVisible(False)
        self.status_label = QLabel("就绪")
        self.status_label.setObjectName("hint")
        self.cache_label = QLabel("")
        self.cache_label.setObjectName("hint")
        bottom.addWidget(self.status_label, 1)
        bottom.addWidget(self.cache_label)
        bottom.addWidget(self.progress)
        root.addLayout(bottom)
        self.setCentralWidget(central)

        self.match_combo.currentIndexChanged.connect(self._on_match_mode_changed)

        # 快捷键
        for seq, fn in (
            ("Ctrl+F", lambda: self.pattern_combo.setFocus()),
            ("Meta+F", lambda: self.pattern_combo.setFocus()),
            ("Ctrl+L", lambda: self.path_edit.setFocus()),
            ("Meta+L", lambda: self.path_edit.setFocus()),
            ("F3", self.next_hit),
            ("Shift+F3", self.prev_hit),
            ("Meta+Shift+R", self.reveal_in_finder),
            ("Ctrl+Shift+R", self.reveal_in_finder),
            ("Meta+O", self.open_external),
            ("Ctrl+O", self.open_external),
            ("Escape", self.stop_search),
            ("Meta+1", lambda: self.left_tabs.setCurrentIndex(0)),
            ("Meta+2", lambda: self.left_tabs.setCurrentIndex(1)),
            ("Meta+3", lambda: self.right_tabs.setCurrentIndex(0)),
            ("Meta+4", lambda: self.right_tabs.setCurrentIndex(1)),
        ):
            QShortcut(QKeySequence(seq), self, activated=fn)

        self._build_menu()
        self.statusBar().showMessage("就绪：选择文件夹，输入正则，开始搜索；或切到「文件」页直接浏览", 5000)

    def _build_menu(self) -> None:
        bar = self.menuBar()
        m_file = bar.addMenu("文件")
        a1 = QAction("在 Finder 中显示", self)
        a1.setShortcut(QKeySequence("Meta+Shift+R"))
        a1.triggered.connect(self.reveal_in_finder)
        m_file.addAction(a1)
        a2 = QAction("用默认程序打开", self)
        a2.triggered.connect(self.open_external)
        m_file.addAction(a2)
        m_file.addSeparator()
        a3 = QAction("导出结果 CSV…", self)
        a3.triggered.connect(self.export_csv)
        m_file.addAction(a3)
        a4 = QAction("导出结果 JSON…", self)
        a4.triggered.connect(self.export_json)
        m_file.addAction(a4)

        m_tools = bar.addMenu("工具")
        a5 = QAction("清理解析缓存", self)
        a5.triggered.connect(self.clear_cache)
        m_tools.addAction(a5)
        a6 = QAction("支持的文件类型…", self)
        a6.triggered.connect(self.show_types)
        m_tools.addAction(a6)
        a_reset = QAction("重置搜索选项", self)
        a_reset.triggered.connect(self.reset_options)
        m_tools.addAction(a_reset)
        a7 = QAction("打开 SQLite 数据库…", self)
        a7.triggered.connect(self.open_db_file)
        m_tools.addAction(a7)

        m_help = bar.addMenu("帮助")
        a8 = QAction("关于 / 后端信息", self)
        a8.triggered.connect(self.show_about)
        m_help.addAction(a8)

    # -- 设置 -----------------------------------------------------------
    def _restore_settings(self) -> None:
        s = self.settings
        path = s.value("path", "", type=str) or os.getcwd()
        self.path_edit.setText(path)
        self.pattern_combo.addItems(s.value("history", [], type=list) or [])
        if self.pattern_combo.count():
            self.pattern_combo.setCurrentIndex(0)
        self.case_combo.setCurrentIndex(max(0, min(s.value("case_mode", 0, type=int), 2)))
        self.match_combo.setCurrentIndex(max(0, min(s.value("match_mode", 0, type=int), 2)))
        self.cb_word.setChecked(False)
        self.cb_deep.setChecked(False)
        self.cb_arch.setChecked(s.value("archives", True, type=bool))
        self.cb_cache.setChecked(s.value("cache", True, type=bool))
        self.type_combo.setCurrentText(s.value("type", "全部", type=str) or "全部")
        self.backend_combo.setCurrentIndex(max(0, min(s.value("backend", 0, type=int), 2)))
        self.workers_spin.setValue(s.value("workers", 0, type=int))
        geo = s.value("geometry")
        if geo:
            self.restoreGeometry(geo)
        self.left_tabs.setCurrentIndex(s.value("left_tab", 0, type=int))
        if os.path.isdir(self.path_edit.text()):
            self.browser.set_root(self.path_edit.text())

    def _save_settings(self) -> None:
        s = self.settings
        s.setValue("path", self.path_edit.text())
        hist = [self.pattern_combo.itemText(i) for i in range(self.pattern_combo.count())][:30]
        s.setValue("history", hist)
        s.setValue("case_mode", self.case_combo.currentIndex())
        s.setValue("match_mode", self.match_combo.currentIndex())
        # 「整词 / 二进制字符串」不持久化：每次启动都从关闭开始
        s.setValue("archives", self.cb_arch.isChecked())
        s.setValue("cache", self.cb_cache.isChecked())
        s.setValue("type", self.type_combo.currentText())
        s.setValue("backend", self.backend_combo.currentIndex())
        s.setValue("workers", self.workers_spin.value())
        s.setValue("geometry", self.saveGeometry())
        s.setValue("left_tab", self.left_tabs.currentIndex())

    def closeEvent(self, event) -> None:
        self._save_settings()
        self.stop_search()
        th = self.search_thread
        if th is not None and th.isRunning():
            th.wait(3000)
        for sc in list(getattr(self.browser, "_scanners", {}).values()):
            try:
                sc.wait(300)
            except Exception:
                pass
        super().closeEvent(event)

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if p:
                self.path_edit.setText(p)
                self.on_path_changed()
                break

    # -- 路径 -----------------------------------------------------------
    def on_path_changed(self) -> None:
        p = self.path_edit.text().strip()
        if os.path.isdir(p):
            self.browser.set_root(p)
        elif os.path.isfile(p):
            self.browser.set_root(os.path.dirname(p))

    def browse_folder(self) -> None:
        start = self.path_edit.text() or os.path.expanduser("~")
        d = QFileDialog.getExistingDirectory(self, "选择文件夹", start)
        if d:
            self.path_edit.setText(d)
            self.on_path_changed()

    def open_db_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "打开 SQLite 数据库", self.path_edit.text() or os.path.expanduser("~"),
            "SQLite (*.db *.sqlite *.sqlite3 *.db3);;所有文件 (*)")
        if path:
            self.right_tabs.setCurrentIndex(1)
            self.db_panel.open_db(path)

    # -- 搜索 -----------------------------------------------------------
    def _exts(self) -> Optional[List[str]]:
        group = KIND_GROUPS.get(self.type_combo.currentText(), ())
        return list(group) or None

    def build_cfg(self) -> SearchConfig:
        p = self.path_edit.text().strip()
        return SearchConfig(
            pattern=self.pattern_combo.currentText(),
            roots=[p] if p else [],
            ignore_case=self.case_combo.currentData() == "ignore",
            case_mode=self.case_combo.currentData() or "smart",
            whole_word=self.cb_word.isChecked(),
            match_mode=self.match_combo.currentData() or "smart",
            multiline=True, dotall=True,
            whitelist_exts=self._exts(),
            exclude_globs=list(DEFAULT_EXCLUDES),
            max_hits_per_file=self.hits_spin.value(),
            context=60,
            backend=self.backend_combo.currentData() or "auto",
            workers=self.workers_spin.value(),
            use_cache=self.cb_cache.isChecked(),
            deep_binary=self.cb_deep.isChecked(),
            archives=self.cb_arch.isChecked(),
            keep_docs=False,
            extract=ExtractContext(
                zip_passwords=[p.strip() for p in self.zip_pwd_edit.text().split(",") if p.strip()],
            ),
        )

    def start_search(self) -> None:
        if self.search_thread is not None and self.search_thread.isRunning():
            return
        cfg = self.build_cfg()
        if not cfg.pattern:
            self._flash("请输入正则表达式")
            return
        if not cfg.roots or not os.path.exists(cfg.roots[0]):
            self._flash("请选择存在的文件夹")
            return
        try:
            SearchEngine(cfg)
        except Exception as e:
            self.pattern_combo.setStyleSheet("border:1px solid #d93025;")
            self._flash(f"正则无效：{e}")
            return
        if cfg.match_mode == "text" and looks_like_regex(cfg.pattern):
            self.pattern_combo.setStyleSheet("border:1px solid #d9a03a;")
            self.statusBar().showMessage(
                "当前是「纯文字」模式：输入里的 (? [ { 会按原样查找。"
                "若想用正则，请把匹配方式切成「智能匹配」或「正则」。", 15000)
        else:
            self.pattern_combo.setStyleSheet("")
        self.zip_pwd_edit.setStyleSheet("")
        self._cfg_last = cfg
        self._remember_pattern(cfg.pattern)

        self.tree.clear()
        self._file_hits.clear()
        self._hit_order.clear()
        self._hit_index = -1
        self._pending.clear()
        self.current_hit = None
        self.left_tabs.setCurrentIndex(0)
        self.btn_search.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.status_label.setText("正在枚举文件…")

        self.search_thread = SearchThread(cfg, self)
        self.search_thread.fileFound.connect(self._queue_file)
        self.search_thread.progress.connect(self._on_progress)
        self.search_thread.finishedReport.connect(self._on_finished)
        self.search_thread.logLine.connect(lambda s: self.statusBar().showMessage(s, 5000))
        self._encrypted_count = 0
        self.search_thread.encryptedFound.connect(self._on_encrypted)
        self.search_thread.start()

    def _remember_pattern(self, pattern: str) -> None:
        if not pattern:
            return
        for i in range(self.pattern_combo.count()):
            if self.pattern_combo.itemText(i) == pattern:
                self.pattern_combo.removeItem(i)
                break
        self.pattern_combo.insertItem(0, pattern)
        self.pattern_combo.setCurrentIndex(0)

    def stop_search(self) -> None:
        if self.search_thread is not None and self.search_thread.isRunning():
            self.search_thread.stop()
            self.status_label.setText("正在停止…")

    def _queue_file(self, fr: FileResult) -> None:
        self._pending.append(fr)
        if not self._flush_timer.isActive():
            self._flush_timer.start()

    def _flush_pending(self) -> None:
        if not self._pending:
            self._flush_timer.stop()
            return
        batch = self._pending[:200]
        self._pending = self._pending[200:]
        self.tree.setUpdatesEnabled(False)
        try:
            first = self.tree.topLevelItemCount() == 0
            for fr in batch:
                self._add_file_item(fr)
            if first and self.tree.topLevelItemCount():
                self.tree.topLevelItem(0).setExpanded(True)
        finally:
            self.tree.setUpdatesEnabled(True)
        if not self._pending:
            self._flush_timer.stop()

    def _add_file_item(self, fr: FileResult) -> None:
        # 压缩包成员显示 member 相对路径（如 inner/报告.docx），普通文件显示相对路径，
        # 这样包内包外的同名文件也能区分；完整路径在 tooltip 里
        short = fr.display_name or os.path.basename(fr.path)
        top = QTreeWidgetItem([f"{short}   ({len(fr.hits)})", "", "",
                               f"{fr.kind} · {fr.chars:,} 字符"])
        top.setData(0, Qt.ItemDataRole.UserRole, fr)
        top.setToolTip(0, fr.path.replace("::", "  ›  "))
        top.setToolTip(3, f"{fr.path}\n字符数 {fr.chars} · 用时 {fr.elapsed*1000:.0f}ms"
                        + ("  [缓存]" if fr.cached else ""))
        font = top.font(0)
        font.setBold(True)
        top.setFont(0, font)
        self._file_hits[fr.path] = fr.hits
        MAX_TREE_HITS = 3000
        for h in fr.hits:
            self._hit_order.append(h)
        for h in fr.hits[:MAX_TREE_HITS]:
            extra = ""
            cells = (h.meta or {}).get("cells")
            if cells:
                for addr, s, e in cells:
                    if s <= h.col < e:
                        extra = f" · {addr}"
                        break
            child = QTreeWidgetItem(["", f"{h.label}{extra}", str(h.line_no),
                                     h.line_text.strip()])
            child.setData(0, Qt.ItemDataRole.UserRole + 1, h)
            child.setData(3, Qt.ItemDataRole.UserRole + 1, h)
            tip = f"{fr.display_name} › {h.label}"
            if "行" not in h.label:
                tip += f" · 第 {h.line_no} 行"
            child.setToolTip(1, tip)
            top.addChild(child)
        if len(fr.hits) > MAX_TREE_HITS:
            more = QTreeWidgetItem([f"… 另有 {len(fr.hits) - MAX_TREE_HITS} 条命中未在树中列出",
                                    "", "", "可用 F3 逐个跳转，或导出 CSV/JSON 获取全部"])
            more.setFlags(Qt.ItemFlag.NoItemFlags)
            more.setForeground(3, QColor("#8a94a0"))
            top.addChild(more)
        self.tree.addTopLevelItem(top)

    def _on_progress(self, done: int, total: int, hits: int, cached: int) -> None:
        if total <= 0:
            return
        if self.progress.maximum() != total:
            self.progress.setRange(0, total)
        self.progress.setValue(done)
        self.status_label.setText(f"已扫描 {done}/{total} 个文件 · 命中 {hits} 处")
        if cached:
            self.cache_label.setText(f"缓存命中 {cached}")

    def _on_finished(self, report: SearchReport) -> None:
        self._flush_pending()
        self.btn_search.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.progress.setVisible(False)
        backend = {"process": "多进程", "thread": "多线程"}.get(report.backend, report.backend)
        txt = (f"完成 · {report.files_matched} 个文件 / {report.total_hits} 处命中 · "
               f"扫描 {report.files_scanned}/{report.total_files} · {report.elapsed:.1f}s")
        if report.cached_files:
            txt += f" · 缓存 {report.cached_files}"
        if backend:
            txt += f" · {backend}"
        if getattr(report, "encrypted", 0):
            txt += f" · {report.encrypted} 个加密已跳过"
            self.btn_retry_enc.setText(f"重搜加密项 ({report.encrypted})")
            self.btn_retry_enc.setVisible(True)
            if not self.zip_pwd_edit.text().strip():
                self.zip_pwd_edit.setFocus()
                self.zip_pwd_edit.setStyleSheet("border:1px solid #d9a03a;")
                self.statusBar().showMessage(
                    "有加密压缩包未读取：在「zip 口令」里填密码后回车即可重搜", 12000)
        else:
            self.btn_retry_enc.setVisible(False)
        if report.cancelled:
            txt = "已停止 · " + txt
        self.status_label.setText(txt)
        self.cache_label.setText("")
        self._update_cache_label()
        self.search_thread = None
        self.statusBar().clearMessage()
        self.apply_filter(self.filter_edit.text())
        if self.tree.topLevelItemCount():
            self.tree.sortItems(0, Qt.SortOrder.AscendingOrder)
            top = self.tree.topLevelItem(0)
            if top.childCount():
                self.tree.setCurrentItem(top.child(0))
                self.show_hit(top.child(0).data(0, Qt.ItemDataRole.UserRole + 1))

    def _on_match_mode_changed(self, _idx: int = 0) -> None:
        mode = self.match_combo.currentData()
        self.pattern_combo.setStyleSheet("")
        if mode == "text":
            self.pattern_combo.lineEdit().setPlaceholderText(
                "纯文字匹配：特殊符号（(?<! [ ] { } | \\）按原样查找")
        elif mode == "regex":
            self.pattern_combo.lineEdit().setPlaceholderText(
                r"正则表达式，例如  张三|138\d{8}    或    (?i)password\s*=")
        else:
            self.pattern_combo.lineEdit().setPlaceholderText(
                "直接输入要查找的文字，也可以写正则——会自动判断")

    def reset_options(self) -> None:
        """一键把容易误触的开关恢复到默认。"""
        self.cb_word.setChecked(False)
        self.cb_deep.setChecked(False)
        self.cb_arch.setChecked(True)
        self.cb_cache.setChecked(True)
        self.case_combo.setCurrentIndex(0)
        self.match_combo.setCurrentIndex(0)
        self.type_combo.setCurrentIndex(0)
        self.backend_combo.setCurrentIndex(0)
        self.workers_spin.setValue(0)
        self.hits_spin.setValue(0)
        self.statusBar().showMessage("搜索选项已重置为默认", 5000)

    def _on_encrypted(self, fr) -> None:
        self._encrypted_count = getattr(self, "_encrypted_count", 0) + 1
        name = getattr(fr, "display_name", "")
        self.statusBar().showMessage(f"{name}：压缩包已加密，需要口令", 4000)

    def _flash(self, msg: str) -> None:
        self.statusBar().showMessage(msg, 4000)
        self.status_label.setText(msg)

    # -- 预览 -----------------------------------------------------------
    def browse_file(self, path: str) -> None:
        """在「文件」页选中或双击文件时预览。"""
        if not os.path.isfile(path):
            return
        self._browse_mode = True
        self.current_hit = None
        self._last_browse = path
        self._open_doc(path, goto_hit=None)

    def open_hit_in_db(self, hit) -> None:
        """双击命中：若来自数据库，切到「数据库」页并选中对应行。"""
        self.show_hit(hit)
        meta = hit.meta or {}
        table = meta.get("table")
        path = self._fs_path(hit.path)
        if not table or not path:
            return
        if os.path.splitext(path)[1].lower() not in DB_SUFFIXES:
            return
        if not os.path.exists(path):
            return
        self.right_tabs.setCurrentIndex(1)
        if self.db_panel.path != path and not self.db_panel.open_db(path):
            return
        self.db_panel.locate(table, meta.get("rowid"), meta.get("row"))

    def show_hit(self, hit) -> None:
        if hit is None:
            return
        self._browse_mode = False
        self.current_hit = hit
        try:
            self._hit_index = self._hit_order.index(hit)
        except ValueError:
            self._hit_index = -1
        if self.current_doc is not None and self.current_path == hit.path:
            self._render_current()
            return
        self._open_doc(hit.path, goto_hit=hit)

    def _open_doc(self, path: str, goto_hit=None) -> None:
        self.preview_title.setText(f"{os.path.basename(path)} — 正在解析…")
        self.loader_token += 1
        token = self.loader_token
        self.loader = DocLoader(path, self._cfg_last, token, self)
        self.loader.loaded.connect(self._on_doc_loaded)
        self.loader.start()

    @Slot(object, int)
    def _on_doc_loaded(self, doc, token: int) -> None:
        if token != self.loader_token:
            return
        if isinstance(doc, Exception):
            self.preview_title.setText("解析失败")
            self.editor.setPlainText(str(doc))
            return
        self.current_doc = doc
        self.current_path = getattr(doc, "path", None)
        self._full_text = doc.text or ""
        self._loaded_chars = 0
        self._doc_terms = []
        self._doc_match_lines = []
        self._doc_match_index = 0
        self._update_chips()
        self._update_filter_info()
        # 数据库：自动装载查询面板
        if self.current_path and os.path.splitext(self.current_path)[1].lower() in DB_SUFFIXES:
            if self.db_panel.path != self.current_path:
                self.db_panel.open_db(self.current_path)
            if self._browse_mode:
                self.right_tabs.setCurrentIndex(1)
        self._render_current()

    def _render_current(self) -> None:
        doc = self.current_doc
        if doc is None:
            return
        if self._doc_terms:
            self._render_filtered(None)
            return
        hit = self.current_hit
        text = self._full_text
        self.editor.blockSignals(True)
        try:
            self.editor.clear()
            if hit is not None:
                self._render_windowed(doc, hit, text)
            else:
                self._render_from_start(text)
        finally:
            self.editor.blockSignals(False)
        if hit is not None:
            pos = f"{hit.display_name} › {hit.label}"
            if "行" not in hit.label:
                pos += f" · 第 {hit.line_no} 行"
            self.preview_title.setText(pos)
            self._render_detail(doc, hit)
            QTimer.singleShot(0, self._center_on_hit)
        else:
            info = f"{os.path.basename(doc.path)} · {doc.kind} · {len(text):,} 字符"
            if doc.error:
                info += f" · 解析提示: {doc.error}"
            elif doc.warning:
                info += f" · {doc.warning}"
            self.preview_title.setText(info)
            self.detail.setHtml(
                f"<b>{self._esc(os.path.basename(doc.path))}</b> — {self._esc(doc.path)}<br>"
                f"类型 {doc.kind} · 结构单元 {len(doc.segments)} 个 · 大小 {human_size(doc.size)}"
                + (f"<br><span style='color:#c5221f'>{self._esc(doc.error)}</span>" if doc.error else "")
                + (f"<br><span style='color:#b06000'>{self._esc(doc.warning)}</span>" if doc.warning else ""))

    def _render_windowed(self, doc, hit, text: str) -> None:
        self._windowed = True
        WIN = self.PREVIEW_WINDOW
        if len(text) <= WIN:
            base, window = 0, text
            line_offset = 0
        else:
            base = max(0, min(hit.start - WIN // 3, len(text) - WIN))
            if base > 0:                                  # 对齐到行首，行号才连续
                nl = text.rfind("\n", 0, base)
                base = nl + 1 if nl >= 0 else base
            window = text[base:base + WIN]
            line_offset = doc.line_of(base) - 1
        self.editor.set_line_offset(line_offset)
        window = safe_text(window)
        self._loaded_chars = base + len(window)
        self._window_start = base
        self.editor.setPlainText(window)
        self._highlight(hit, base, window)

    def _render_from_start(self, text: str) -> None:
        self._windowed = False
        self.editor.set_line_offset(0)
        chunk = safe_text(text[: self.PREVIEW_CHUNK])
        self.editor.setPlainText(chunk)
        self._loaded_chars = self.PREVIEW_CHUNK if len(text) > self.PREVIEW_CHUNK else len(text)
        self.editor.setExtraSelections([])
        cur = QTextCursor(self.editor.document())
        cur.movePosition(QTextCursor.MoveOperation.Start)
        self.editor.setTextCursor(cur)
        self.editor.verticalScrollBar().setValue(0)
        if self._loaded_chars < len(text):
            self.statusBar().showMessage(
                f"已加载 {self._loaded_chars:,} / {len(text):,} 字符，滚动到底部继续加载", 6000)

    def _load_more(self) -> None:
        if getattr(self, "_windowed", False):
            return
        text = self._full_text
        if not text or self._loaded_chars >= len(text):
            return
        end = min(len(text), self._loaded_chars + self.PREVIEW_CHUNK)
        chunk = safe_text(text[self._loaded_chars:end])
        self.editor.append_chunk(chunk)
        self._loaded_chars = end
        if self._loaded_chars >= len(text):
            self.statusBar().showMessage("已到文件末尾", 3000)

    def _highlight(self, hit, base: int, window: str) -> None:
        sels: List[QTextEdit.ExtraSelection] = []
        for h in self._file_hits.get(hit.path, []):
            if h.end <= base or h.start >= base + len(window):
                continue
            sel = QTextEdit.ExtraSelection()
            fmt = QTextCharFormat()
            fmt.setBackground(HIT_BG)
            fmt.setForeground(HIT_FG)
            sel.format = fmt
            cur = QTextCursor(self.editor.document())
            cur.setPosition(max(0, h.start - base))
            cur.setPosition(max(0, min(len(window), h.end - base)),
                            QTextCursor.MoveMode.KeepAnchor)
            sel.cursor = cur
            sels.append(sel)
        start = max(0, hit.start - base)
        cur = QTextCursor(self.editor.document())
        cur.setPosition(start)
        cur.setPosition(max(start + 1, min(len(window), hit.end - base)),
                        QTextCursor.MoveMode.KeepAnchor)
        cur_sel = QTextEdit.ExtraSelection()
        fmt = QTextCharFormat()
        fmt.setBackground(CUR_BG)
        fmt.setForeground(CUR_FG)
        fmt.setFontWeight(QFont.Weight.Bold)
        cur_sel.format = fmt
        cur_sel.cursor = cur
        sels.append(cur_sel)
        self.editor.setExtraSelections(sels)
        nav = QTextCursor(self.editor.document())
        nav.setPosition(start)
        self.editor.setTextCursor(nav)
        self.editor.ensureCursorVisible()

    def _center_on_hit(self) -> None:
        try:
            self.editor.centerCursor()
        except Exception:
            pass

    # -- 文件内继续查找（可叠加多个条件） -------------------------------
    def add_doc_filter(self, text=None, exclude: Optional[bool] = None) -> None:
        if self.current_doc is None:
            self._flash("请先打开一个文件")
            return
        raw = (text if isinstance(text, str) else self.doc_filter_edit.text()).strip()
        if not raw:
            return
        if exclude is None:
            exclude = self.filter_mode_combo.currentData() == "exclude"
            if raw[0] in "-!" and len(raw) > 1:
                exclude = True
                raw = raw[1:].strip()
        if not raw:
            return
        item = (raw, bool(exclude))
        if item in self._doc_terms:
            self.doc_filter_edit.clear()
            return
        self._doc_terms.append(item)
        self.doc_filter_edit.clear()
        self._apply_doc_filter_change()

    def add_doc_filter_exclude(self, text: str) -> None:
        self.add_doc_filter(text, exclude=True)

    def remove_doc_filter_at(self, index: int) -> None:
        if 0 <= index < len(self._doc_terms):
            del self._doc_terms[index]
            self._apply_doc_filter_change()

    def _apply_doc_filter_change(self) -> None:
        self._update_chips()
        self._compute_doc_matches()
        self._doc_match_index = 0
        if self._doc_terms:
            center = self._doc_match_lines[0][0] if self._doc_match_lines else None
            self._render_filtered(center)
        else:
            self._render_current()
        self._update_filter_info()

    def clear_doc_filters(self) -> None:
        if not self._doc_terms:
            return
        self._doc_terms = []
        self._doc_match_lines = []
        self._doc_match_index = 0
        self._update_chips()
        self._update_filter_info()
        self._render_current()

    def goto_doc_match(self, step: int) -> None:
        if not self._doc_terms:
            return
        if not self._doc_match_lines:
            self._flash("当前没有满足所有条件的行")
            return
        self._doc_match_index = (self._doc_match_index + step) % len(self._doc_match_lines)
        self._render_filtered(self._doc_match_lines[self._doc_match_index][0])
        self._update_filter_info()

    def _update_chips(self) -> None:
        # 布局尾部固定是 [stretch, filter_info]，只回收中间的 chip
        while self.filter_chips_layout.count() > 2:
            item = self.filter_chips_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        for i, (term, exc) in enumerate(self._doc_terms):
            btn = QToolButton()
            label = term if len(term) <= 18 else term[:16] + "…"
            btn.setText(f"{'− ' if exc else ''}{label}  ×")
            btn.setObjectName("chipExclude" if exc else "chip")
            btn.setToolTip(f"{'排除' if exc else '包含'}：{term}（点击移除）")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda _checked=False, idx=i: self.remove_doc_filter_at(idx))
            self.filter_chips_layout.insertWidget(self.filter_chips_layout.count() - 2, btn)


    def _update_filter_info(self) -> None:
        if not self._doc_terms:
            self.filter_info.setText("")
            return
        n = len(self._doc_match_lines)
        if n:
            self.filter_info.setText(f"{n} 行匹配 · 第 {self._doc_match_index + 1}/{n}")
        else:
            self.filter_info.setText("0 行匹配")

    def _compute_doc_matches(self) -> None:
        self._doc_match_lines = []
        if not self._doc_terms or not self._full_text:
            return
        ci = (self.case_combo.currentData() or "smart") != "sensitive"
        hay = self._full_text.lower() if ci else self._full_text
        norm = (lambda x: x.lower()) if ci else (lambda x: x)
        inc = [norm(t) for t, ex in self._doc_terms if not ex]
        exc = [norm(t) for t, ex in self._doc_terms if ex]
        pos = 0
        n = len(hay)
        cap = 20000
        while pos < n and len(self._doc_match_lines) < cap:
            nl = hay.find("\n", pos)
            end = n if nl < 0 else nl
            line = hay[pos:end]
            if all(nd in line for nd in inc) and not any(nd in line for nd in exc):
                self._doc_match_lines.append((pos, end))
            pos = end + 1

    def _window_slice(self, center: int):
        text = self._full_text
        WIN = self.PREVIEW_WINDOW
        if len(text) <= WIN:
            return 0, text, 0
        base = max(0, min(center - WIN // 3, len(text) - WIN))
        if base > 0:
            nl = text.rfind("\n", 0, base)
            base = nl + 1 if nl >= 0 else base
        window = text[base:base + WIN]
        offset = self.current_doc.line_of(base) - 1 if self.current_doc else 0
        return base, window, offset

    def _render_filtered(self, center: Optional[int]) -> None:
        doc = self.current_doc
        if doc is None:
            return
        if center is None:
            center = self._doc_match_lines[0][0] if self._doc_match_lines else 0
        self._windowed = len(self._full_text) > self.PREVIEW_WINDOW
        base, window, line_offset = self._window_slice(center)
        self._window_base = base
        self._win_len = len(window)
        self.editor.blockSignals(True)
        try:
            self.editor.setPlainText(safe_text(window))
        finally:
            self.editor.blockSignals(False)
        self.editor.set_line_offset(line_offset)
        self._apply_filter_selections()
        pos = max(0, min(len(window), center - base))
        nav = QTextCursor(self.editor.document())
        nav.setPosition(pos)
        self.editor.setTextCursor(nav)
        self.editor.ensureCursorVisible()
        QTimer.singleShot(0, self._center_on_hit)
        inc_list = [t for t, ex in self._doc_terms if not ex]
        exc_list = [t for t, ex in self._doc_terms if ex]
        parts = []
        if inc_list:
            parts.append(" + ".join(inc_list))
        if exc_list:
            parts.append("排除 " + " / ".join(exc_list))
        cond = "；".join(parts)
        if self._doc_match_lines:
            title = f"{os.path.basename(doc.path)} · 文件内筛选：{cond}"
        else:
            title = f"{os.path.basename(doc.path)} · 没有满足「{cond}」的行（下方为全文）"
        self.preview_title.setText(title)

    def _apply_filter_selections(self) -> None:
        window = self.editor.toPlainText()
        if not window:
            return
        base = self._window_base
        wend = base + len(window)
        ci = (self.case_combo.currentData() or "smart") != "sensitive"
        low = window.lower() if ci else window
        needles = [t.lower() if ci else t for t, ex in self._doc_terms if not ex]
        cur = None
        if 0 <= self._doc_match_index < len(self._doc_match_lines):
            cur = self._doc_match_lines[self._doc_match_index][0]
        sels: List[QTextEdit.ExtraSelection] = []

        def add(a: int, b: int, bg: QColor, fg: Optional[QColor] = None) -> None:
            if b <= a:
                return
            sel = QTextEdit.ExtraSelection()
            fmt = QTextCharFormat()
            fmt.setBackground(bg)
            if fg is not None:
                fmt.setForeground(fg)
            cur_cur = QTextCursor(self.editor.document())
            cur_cur.setPosition(max(0, a))
            cur_cur.setPosition(max(a, b), QTextCursor.MoveMode.KeepAnchor)
            sel.format = fmt
            sel.cursor = cur_cur
            sels.append(sel)

        # 1) 命中行整行底色
        shown = 0
        for ls, le in self._doc_match_lines:
            if le <= base or ls >= wend:
                continue
            shown += 1
            if shown > 4000:
                break
            a = max(0, ls - base)
            b = min(len(window), le - base)
            add(a, b, QColor("#fff2cc") if cur != ls else QColor("#ffd28a"))
        # 2) 关键词高亮
        for ls, le in self._doc_match_lines:
            if le <= base or ls >= wend:
                continue
            a = max(0, ls - base)
            b = min(len(window), le - base)
            line_low = low[a:b]
            for nd in needles:
                p = 0
                while True:
                    j = line_low.find(nd, p)
                    if j < 0:
                        break
                    add(a + j, a + j + len(nd), QColor("#ffca28"), QColor("#4a3200"))
                    p = j + len(nd)
        self.editor.setExtraSelections(sels)

    def _render_detail(self, doc, hit) -> None:
        meta = hit.meta or {}
        rows: List[str] = []
        rows.append(f"<b>位置</b> {self._esc(hit.label)}")
        if "行" not in hit.label:
            rows.append(f" &nbsp; <b>行</b> {hit.line_no}")
        rows.append(f" &nbsp; <b>列</b> {hit.col + 1} &nbsp; <b>偏移</b> {hit.start}-{hit.end}<br>")
        if meta.get("table"):
            rows.append(f"<b>数据库表</b> {self._esc(str(meta.get('table')))} "
                        f"<b>行</b> {meta.get('row')} &nbsp;")
        if meta.get("sheet"):
            rows.append(f"<b>工作表</b> {self._esc(str(meta.get('sheet')))} "
                        f"<b>行</b> {meta.get('row')} &nbsp;")
        if meta.get("page"):
            rows.append(f"<b>页码</b> {meta.get('page')} &nbsp;")
        if meta.get("slide"):
            rows.append(f"<b>幻灯片</b> {meta.get('slide')} &nbsp;")
        if meta.get("member"):
            rows.append(f"<b>压缩包</b> {self._esc(str(meta.get('archive_name') or ''))} "
                        f"&nbsp; <b>成员</b> {self._esc(str(meta.get('member')))}<br>")
        ctx = (f"<span style='color:#888'>{self._esc(hit.before)}</span>"
               f"<span style='background:#fdf0cb;color:#4a3f18'>{self._esc(hit.matched)}</span>"
               f"<span style='color:#888'>{self._esc(hit.after)}</span>")
        cols = meta.get("columns") or {}
        table_html = ""
        if cols:
            cells = "".join(
                f"<tr><td style='padding:1px 8px;color:#666'>{self._esc(str(k))}</td>"
                f"<td style='padding:1px 8px'>{self._esc(str(v))}</td></tr>"
                for k, v in list(cols.items())[:10])
            table_html = f"<table cellspacing=0 cellpadding=0>{cells}</table>"
        self.detail.setHtml("".join(rows) + f"<div style='margin-top:3px;font-family:monospace'>{ctx}</div>" + table_html)

    @staticmethod
    def _esc(s: str) -> str:
        return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    # -- 导航 -----------------------------------------------------------
    def next_hit(self) -> None:
        self._navigate(1)

    def prev_hit(self) -> None:
        self._navigate(-1)

    def _navigate(self, step: int) -> None:
        if not self._hit_order:
            return
        idx = self._hit_index if self._hit_index >= 0 else 0
        idx = (idx + step) % len(self._hit_order)
        hit = self._hit_order[idx]
        self._hit_index = idx
        self.show_hit(hit)
        item = self._find_item(hit)
        if item is not None:
            self.tree.setCurrentItem(item)
            self.tree.scrollToItem(item)

    def _find_item(self, hit):
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            fr = top.data(0, Qt.ItemDataRole.UserRole)
            if fr is not None and getattr(fr, "path", None) == hit.path:
                for j in range(top.childCount()):
                    ch = top.child(j)
                    if ch.data(0, Qt.ItemDataRole.UserRole + 1) is hit:
                        return ch
        return None

    # -- 外部 -----------------------------------------------------------
    def _current_target_path(self) -> Optional[str]:
        if self.current_hit is not None:
            return self.current_hit.path
        if self._browse_mode and self._last_browse:
            return self._last_browse
        item = self.tree.currentItem()
        if item is not None:
            fr = item.data(0, Qt.ItemDataRole.UserRole)
            if fr is not None:
                return getattr(fr, "path", None)
            top = item.parent()
            if top is not None:
                fr = top.data(0, Qt.ItemDataRole.UserRole)
                if fr is not None:
                    return getattr(fr, "path", None)
        return self.current_path

    @staticmethod
    def _fs_path(path: Optional[str]) -> Optional[str]:
        """把 "a.zip::member" 还原成磁盘路径。"""
        if not path:
            return path
        return path.split("::", 1)[0]

    def open_external(self) -> None:
        path = self._fs_path(self._current_target_path())
        if not path or not os.path.exists(path):
            return
        subprocess.Popen(["open", path] if sys.platform == "darwin" else ["xdg-open", path])

    def reveal_in_finder(self) -> None:
        path = self._fs_path(self._current_target_path())
        if not path or not os.path.exists(path):
            return
        if sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path])
        elif os.name == "nt":
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(path)])

    def export_csv(self) -> None:
        if not self._hit_order:
            self._flash("没有可导出的结果")
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出结果",
                                              os.path.expanduser("~/filegrep_hits.csv"),
                                              "CSV (*.csv)")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["文件", "路径", "类型", "位置", "行号", "列", "匹配", "上下文"])
            for h in self._hit_order:
                w.writerow([h.display_name, h.path, h.kind, h.label, h.line_no, h.col + 1,
                            h.matched, h.line_text.strip()])
        self.statusBar().showMessage(f"已导出 {len(self._hit_order)} 条到 {path}", 6000)

    def export_json(self) -> None:
        if not self._hit_order:
            self._flash("没有可导出的结果")
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出结果",
                                              os.path.expanduser("~/filegrep_hits.json"),
                                              "JSON (*.json)")
        if not path:
            return
        data = [{"file": h.display_name, "path": h.path, "kind": h.kind, "label": h.label,
                 "line": h.line_no, "col": h.col + 1, "match": h.matched,
                 "text": h.line_text.strip(), "meta": h.meta} for h in self._hit_order]
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.statusBar().showMessage(f"已导出 {len(data)} 条到 {path}", 6000)

    def clear_cache(self) -> None:
        from .cache import get_cache
        get_cache(True).clear()
        self._update_cache_label()
        self.statusBar().showMessage("解析缓存已清理", 5000)

    def _update_cache_label(self) -> None:
        try:
            from .cache import get_cache
            st = get_cache(True).stats()
            if st.get("enabled"):
                self.cache_label.setText(
                    f"缓存 {st.get('entries', 0)} 项 / {st.get('bytes', 0)/1048576:.1f}MB")
        except Exception:
            pass

    def show_types(self) -> None:
        suffixes = supported_suffixes()
        QMessageBox.information(self, "支持的文件类型",
                                f"共 {len(suffixes)} 种扩展名：\n\n" + "  ".join(suffixes[:80])
                                + ("\n…" if len(suffixes) > 80 else ""))

    def show_about(self) -> None:
        from . import __version__
        from .parsers import available_backends
        bk = available_backends()
        txt = "\n".join(f"{k}: {v}" for k, v in bk.items())
        QMessageBox.about(self, "关于 FileGrep",
                          f"<b>FileGrep {__version__}</b><br>全格式检索 / 浏览 / 数据库查询"
                          f"<br><br><b>解析后端</b><br><pre>{self._esc(txt)}</pre>"
                          "<br>⌘F 正则 · ⌘L 目录 · F3 下一个 · ⌘Shift+R 在 Finder 显示 · ⌘1/⌘2 切页")

    def apply_filter(self, text: str) -> None:
        """结果列表过滤：空格分隔多个词（AND），以 - 或 ! 开头表示排除。"""
        parts = [p for p in (text or "").split() if p]
        include = [p.lower() for p in parts if not p.startswith(("-", "!"))]
        exclude = [p.lstrip("-!").lower() for p in parts if p.startswith(("-", "!")) and len(p) > 1]
        total = 0
        shown = 0
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            visible = 0
            for j in range(top.childCount()):
                ch = top.child(j)
                total += 1
                hay = ((ch.text(3) or "") + " " + (ch.text(1) or "") + " " + (ch.text(0) or "")).lower()
                show = all(k in hay for k in include) and not any(k in hay for k in exclude)
                ch.setHidden(not show)
                if show:
                    visible += 1
                    shown += 1
            top.setHidden(visible == 0 and bool(parts))
        if not parts:
            # 未过滤时报告真实命中总数（树里每文件只列前 3000 条）
            real = len(self._hit_order)
            self.result_filter_info.setText(f"共 {real} 处" if real else "")
        else:
            self.result_filter_info.setText(f"显示 {shown}/{total} 条")


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def main(argv: Optional[List[str]] = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    app = QApplication.instance() or QApplication(sys.argv[:1])
    app.setApplicationName(APP_NAME)
    app.setApplicationDisplayName(APP_NAME)
    app.setOrganizationName("filegrep")
    win = MainWindow()
    for a in argv:
        if os.path.exists(a):
            win.path_edit.setText(os.path.abspath(a))
            win.on_path_changed()
            break
    win.show()
    return app.exec()
