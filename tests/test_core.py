"""回归测试：python tests/test_core.py"""

from __future__ import annotations

import os
import shutil
import sqlite3
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

import filegrep                                     # noqa: E402
from filegrep import SearchConfig, SearchEngine, load_document   # noqa: E402
from filegrep.cache import DocCache                 # noqa: E402
from filegrep.extract import ExtractContext, Source, extract_source  # noqa: E402

SAMPLES = os.path.join(HERE, "samples")


def ensure_samples() -> None:
    if not os.path.isdir(SAMPLES) or len(os.listdir(SAMPLES)) < 10:
        import importlib.util
        spec = importlib.util.spec_from_file_location("mk", os.path.join(HERE, "make_samples.py"))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        mod.main()


class TestExtractors(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        ensure_samples()

    def _doc(self, name: str):
        return load_document(os.path.join(SAMPLES, name), use_cache=False)

    def test_text_and_gbk(self):
        self.assertIn("张三", self._doc("notes.txt").text)
        gbk = self._doc("legacy_gbk.txt")
        self.assertIn("张三", gbk.text, "GB18030 编码应自动识别")
        self.assertEqual(gbk.meta.get("encoding"), "gb18030")

    def test_csv_rows(self):
        doc = self._doc("data.csv")
        self.assertIn("410105199001011234", doc.text)
        labels = [s.label for s in doc.segments]
        self.assertIn("第 2 行", labels)

    def test_docx_structure(self):
        doc = self._doc("报告.docx")
        self.assertIn("张三", doc.text)
        self.assertTrue(any("表格" in s.label for s in doc.segments), "应识别表格单元格")
        self.assertIn("410105199001011234", doc.text, "表格内容应被抽取")

    def test_xlsx_cells(self):
        doc = self._doc("账本.xlsx")
        self.assertIn("secret_token_9f3a", doc.text)
        seg = next(s for s in doc.segments
                   if s.meta.get("sheet") == "服务器" and s.meta.get("row") == 2)
        self.assertTrue(any(c[0] == "B2" for c in seg.meta.get("cells", [])), "应保留单元格地址")

    def test_pptx_slides(self):
        doc = self._doc("演示.pptx")
        self.assertIn("192.168.1.88", doc.text)
        self.assertTrue(any(s.meta.get("slide") for s in doc.segments))

    def test_pdf_pages(self):
        doc = self._doc("manual.pdf")
        self.assertIn("192.168.1.88", doc.text)
        seg = next(s for s in doc.segments if "secret_token" in doc.text[s.start:s.end])
        self.assertEqual(seg.meta.get("page"), 2)

    def test_sqlite_tables(self):
        doc = self._doc("users.db")
        self.assertIn("admin@example.com", doc.text)
        seg = next(s for s in doc.segments if "admin@example.com" in doc.text[s.start:s.end])
        self.assertEqual(seg.meta.get("table"), "users")
        self.assertEqual(seg.meta.get("row"), 1)
        self.assertIn("name", seg.meta.get("columns", {}))

    def test_eml(self):
        doc = self._doc("mail.eml")
        self.assertIn("secret_token_9f3a", doc.text)
        self.assertTrue(any("正文" in s.label for s in doc.segments))

    def test_rtf(self):
        doc = self._doc("legacy.rtf")
        self.assertIn("secret_token_9f3a", doc.text)

    def test_binary_strings(self):
        doc = self._doc("raw.bin")
        self.assertIn("secret_token_9f3a", doc.text)
        self.assertTrue(any("0x" in s.label for s in doc.segments), "应带文件偏移")

    def test_zip_members(self):
        ctx = ExtractContext(archive_depth=1)
        doc = extract_source(Source(os.path.join(SAMPLES, "bundle.zip")), ctx)
        self.assertIn("压缩包内部", doc.text)
        self.assertTrue(any("inner/" in s.label for s in doc.segments))
        # 内嵌 docx 也要解析
        self.assertIn("张三", doc.text)

    def test_broken_file_never_raises(self):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            f.write(b"not a zip at all")
            bad = f.name
        try:
            doc = extract_source(Source(bad), ExtractContext())
            self.assertIsInstance(doc.text, str)      # 不抛异常即可
        finally:
            os.unlink(bad)


class TestSearch(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        ensure_samples()

    def _run(self, pattern: str, **kw):
        cfg = SearchConfig(pattern=pattern, roots=[SAMPLES], backend="thread",
                           use_cache=False, **kw)
        return cfg, SearchEngine(cfg).run()

    def test_counts_and_positions(self):
        cfg, report = self._run("张三")
        self.assertGreaterEqual(report.files_matched, 6)
        hit = next(h for h in (h for fr in report.results for h in fr.hits)
                   if h.display_name.endswith("data.csv"))
        self.assertEqual(hit.line_no, 2)
        self.assertEqual(hit.label, "第 2 行")

    def test_regex_groups_and_flags(self):
        cfg, report = self._run(r"192\.168\.1\.\d+")
        self.assertGreater(report.total_hits, 3)
        # smart case：pattern 含大写 -> 区分大小写；显式 ignore 才忽略
        cfg_smart = SearchConfig(pattern="SECRET_TOKEN", roots=[SAMPLES], backend="thread",
                                 case_mode="smart", use_cache=False)
        self.assertEqual(SearchEngine(cfg_smart).run().total_hits, 0)
        cfg_ignore = SearchConfig(pattern="SECRET_TOKEN", roots=[SAMPLES], backend="thread",
                                  case_mode="ignore", use_cache=False)
        self.assertGreater(SearchEngine(cfg_ignore).run().total_hits, 0)
        cfg_sens = SearchConfig(pattern="SECRET_TOKEN", roots=[SAMPLES], backend="thread",
                                case_mode="sensitive", use_cache=False)
        self.assertEqual(SearchEngine(cfg_sens).run().total_hits, 0)

    def test_literal_mode(self):
        cfg = SearchConfig(pattern="192.168.1.88", literal=True, roots=[SAMPLES],
                           backend="thread", use_cache=False)
        report = SearchEngine(cfg).run()
        self.assertGreater(report.total_hits, 0)

    def test_extension_filter(self):
        cfg = SearchConfig(pattern="张三", roots=[SAMPLES], backend="thread",
                           whitelist_exts=[".db"], use_cache=False)
        report = SearchEngine(cfg).run()
        self.assertEqual(report.total_files, 1)
        self.assertEqual(report.results[0].kind, "sqlite")

    def test_max_hits_per_file(self):
        cfg = SearchConfig(pattern=".", roots=[SAMPLES], backend="thread",
                           max_hits_per_file=5, use_cache=False)
        report = SearchEngine(cfg).run()
        self.assertTrue(all(len(fr.hits) <= 5 for fr in report.results))

    def test_whole_word(self):
        cfg = SearchConfig(pattern="admin", whole_word=True, roots=[SAMPLES],
                           backend="thread", use_cache=False)
        report = SearchEngine(cfg).run()
        self.assertGreater(report.total_hits, 0)
        for fr in report.results:
            for h in fr.hits:
                self.assertIn(h.matched.lower(), ("admin",))


class TestCache(unittest.TestCase):
    def test_roundtrip_and_invalidation(self):
        tmp = tempfile.mkdtemp()
        try:
            cache = DocCache(tmp, enabled=True)
            path = os.path.join(SAMPLES, "notes.txt")
            doc = load_document(path, use_cache=False)
            st = os.stat(path)
            cache.put(doc, st.st_mtime_ns, "v1")
            got = cache.get(path, st.st_mtime_ns, st.st_size, "v1")
            self.assertIsNotNone(got)
            self.assertEqual(got.text, doc.text)
            self.assertEqual(len(got.segments), len(doc.segments))
            # 变体不同 -> 不命中
            self.assertIsNone(cache.get(path, st.st_mtime_ns, st.st_size, "v2"))
            # mtime 变化 -> 不命中
            self.assertIsNone(cache.get(path, st.st_mtime_ns + 1, st.st_size, "v1"))
            self.assertGreater(cache.stats()["entries"], 0)
            cache.clear()
            self.assertEqual(cache.stats()["entries"], 0)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class TestSqliteReadOnly(unittest.TestCase):
    def test_write_blocked_by_ro_uri(self):
        tmp = tempfile.mkdtemp()
        try:
            db = os.path.join(tmp, "t.db")
            con = sqlite3.connect(db)
            con.execute("CREATE TABLE t(a)")
            con.execute("INSERT INTO t VALUES(1)")
            con.commit()
            con.close()
            from filegrep.extract import _sqlite_open
            ro = _sqlite_open(db)
            try:
                with self.assertRaises(sqlite3.DatabaseError):
                    ro.execute("INSERT INTO t VALUES(2)")
            finally:
                ro.close()
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
