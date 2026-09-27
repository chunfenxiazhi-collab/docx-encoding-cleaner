# -*- coding: utf-8 -*-
"""docx_clean 的单元测试（只用标准库）。"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
import unittest
import uuid
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

# 默认用系统临时目录；沙箱/受限环境可用 DOCX_CLEAN_TEST_TMP 指向一个可写目录
TMP_ROOT = Path(os.environ.get("DOCX_CLEAN_TEST_TMP") or tempfile.gettempdir())
TMP_ROOT.mkdir(parents=True, exist_ok=True)

import docx_clean  # noqa: E402
from tests import fixtures  # noqa: E402


class DocxCleanTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = TMP_ROOT / f"case_{uuid.uuid4().hex[:10]}"
        self.tmp.mkdir(parents=True, exist_ok=True)
        self.src = fixtures.build_docx(self.tmp / "报告.docx")

    def tearDown(self) -> None:
        # ignore_errors：某些受限/沙箱环境不允许删目录，残留无副作用
        shutil.rmtree(self.tmp, ignore_errors=True)

    # ---------- 基本行为 ----------

    def test_default_output_name(self):
        stats = docx_clean.clean_docx(self.src)
        self.assertEqual(stats.output.name, "报告_clean.docx")
        self.assertTrue(stats.output.exists())
        self.assertTrue(self.src.exists(), "原文件必须保留")

    def test_counts_match_fixture(self):
        stats = docx_clean.clean_docx(self.src)
        self.assertEqual(stats.feff, 2, "两个行尾 U+FEFF")
        self.assertEqual(stats.zero_width, 3, "U+200B/U+200D/U+2060 各一个")
        self.assertEqual(stats.cr, 3, "三个 <w:cr/> 软回车")
        self.assertIn("word/document.xml", stats.parts_touched[0])

    def test_junk_is_gone_and_cr_becomes_br(self):
        stats = docx_clean.clean_docx(self.src)
        xml = fixtures.read_part(stats.output, "word/document.xml")
        self.assertNotIn("\ufeff", xml)
        self.assertNotIn("\u200b", xml)
        self.assertNotIn("\u200d", xml)
        self.assertNotIn("\u2060", xml)
        self.assertNotIn("<w:cr/>", xml)
        self.assertEqual(xml.count("<w:br/>"), 3)

    def test_only_junk_changed(self):
        """正文文字一个字都不能变：清理结果 == 对原文做同样的替换。"""
        stats = docx_clean.clean_docx(self.src)
        before = fixtures.read_part(self.src, "word/document.xml")
        expected, *_ = docx_clean._clean_text(before)
        self.assertEqual(fixtures.read_part(stats.output, "word/document.xml"), expected)

    def test_text_content_preserved(self):
        stats = docx_clean.clean_docx(self.src)
        xml = fixtures.read_part(stats.output, "word/document.xml")
        for text in ("人物介绍", "人物 1：甲", "待业青年", "人物 2：EXAMPLE-01", "故事背景"):
            self.assertIn(text, xml)

    # ---------- 结构完整性 ----------

    def test_all_parts_and_order_preserved(self):
        stats = docx_clean.clean_docx(self.src)
        self.assertEqual(fixtures.namelist(self.src), fixtures.namelist(stats.output))

    def test_binary_and_clean_parts_untouched(self):
        stats = docx_clean.clean_docx(self.src)
        self.assertEqual(
            fixtures.read_bytes(self.src, "word/media/image1.png"),
            fixtures.read_bytes(stats.output, "word/media/image1.png"),
        )
        self.assertEqual(
            fixtures.read_part(self.src, "word/styles.xml"),
            fixtures.read_part(stats.output, "word/styles.xml"),
        )
        self.assertEqual(
            fixtures.read_part(self.src, "[Content_Types].xml"),
            fixtures.read_part(stats.output, "[Content_Types].xml"),
        )

    def test_output_is_still_a_readable_zip(self):
        stats = docx_clean.clean_docx(self.src)
        with zipfile.ZipFile(stats.output) as z:
            self.assertIsNone(z.testzip())

    # ---------- 幂等 / 干净文件 ----------

    def test_idempotent(self):
        first = docx_clean.clean_docx(self.src)
        # 显式 -o：即使已经很干净也照写一份规范副本
        second = docx_clean.clean_docx(first.output, self.tmp / "again.docx")
        self.assertEqual(second.junk, 0)
        self.assertFalse(second.parts_touched)
        self.assertEqual(
            fixtures.read_part(first.output, "word/document.xml"),
            fixtures.read_part(second.output, "word/document.xml"),
        )
        # 默认路径：干净文件不产出副本
        third = docx_clean.clean_docx(first.output)
        self.assertEqual(third.junk, 0)
        self.assertIsNone(third.output)
        self.assertFalse((self.tmp / "报告_clean_clean.docx").exists())

    def test_clean_file_reports_nothing(self):
        clean_src = fixtures.build_docx(
            self.tmp / "clean.docx",
            fixtures.DOCUMENT_WITH_JUNK.replace("\ufeff", "")
            .replace("\u200b", "")
            .replace("\u200d", "")
            .replace("\u2060", "")
            .replace("<w:cr/>", "<w:br/>"),
        )
        stats = docx_clean.clean_docx(clean_src)
        self.assertEqual(stats.junk, 0)
        self.assertFalse((self.tmp / "clean_clean.docx").exists(), "干净文件不该产出新文件")

    # ---------- 选项 ----------

    def test_dry_run_writes_nothing(self):
        target = self.tmp / "out.docx"
        stats = docx_clean.clean_docx(self.src, target, dry_run=True)
        self.assertEqual(stats.junk, 8)
        self.assertFalse(target.exists())

    def test_explicit_output_path(self):
        target = self.tmp / "sub" 
        target.mkdir()
        out = target / "fixed.docx"
        stats = docx_clean.clean_docx(self.src, out)
        self.assertEqual(stats.output, out)
        self.assertTrue(out.exists())

    def test_in_place_keeps_backup(self):
        stats = docx_clean.clean_docx(self.src, in_place=True)
        backup = self.src.with_suffix(".docx.bak")
        self.assertTrue(backup.exists())
        self.assertEqual(stats.output, self.src)
        self.assertNotIn("\ufeff", fixtures.read_part(self.src, "word/document.xml"))
        self.assertIn("\ufeff", fixtures.read_part(backup, "word/document.xml"))

    def test_rejects_non_docx(self):
        other = self.tmp / "x.txt"
        other.write_text("hi", encoding="utf-8")
        with self.assertRaises(ValueError):
            docx_clean.clean_docx(other)

    def test_missing_file_raises(self):
        with self.assertRaises(FileNotFoundError):
            docx_clean.clean_docx(self.tmp / "nope.docx")

    # ---------- CLI ----------

    def test_cli_dry_run_exit_code_and_output(self):
        import contextlib
        import io

        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            code = docx_clean.main([str(self.src), "--dry-run"])
        self.assertEqual(code, 0)
        self.assertIn("U+FEFF=2", buf.getvalue())

    def test_cli_json(self):
        import contextlib
        import io
        import json

        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            code = docx_clean.main([str(self.src), "--json"])
        self.assertEqual(code, 0)
        payload = json.loads(buf.getvalue().strip().splitlines()[-1])
        self.assertEqual(payload["junk_total"], 8)


if __name__ == "__main__":
    unittest.main(verbosity=2)
