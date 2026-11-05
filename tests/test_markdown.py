import tempfile
import multiprocessing
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from markdown_backend import convert_pdf_folder_report
from markdown_backend import _convert_pdf
from office_backend import ConversionCancelled


def stalled_markdown_worker(connection, source, staged):
    time.sleep(20)


def create_text_pdf(path):
    from pypdf import PdfWriter
    from pypdf.generic import DictionaryObject, NameObject, DecodedStreamObject, ArrayObject, TextStringObject, NumberObject
    writer = PdfWriter()
    page = writer.add_blank_page(600, 800)
    font = DictionaryObject({NameObject("/Type"): NameObject("/Font"),
                             NameObject("/Subtype"): NameObject("/Type1"),
                             NameObject("/BaseFont"): NameObject("/Helvetica")})
    cid = DictionaryObject({NameObject("/Type"): NameObject("/Font"),
                            NameObject("/Subtype"): NameObject("/CIDFontType0"),
                            NameObject("/BaseFont"): NameObject("/STSong-Light"),
                            NameObject("/CIDSystemInfo"): DictionaryObject({
                                NameObject("/Registry"): TextStringObject("Adobe"),
                                NameObject("/Ordering"): TextStringObject("GB1"),
                                NameObject("/Supplement"): NumberObject(4)})})
    cjk = DictionaryObject({NameObject("/Type"): NameObject("/Font"),
                            NameObject("/Subtype"): NameObject("/Type0"),
                            NameObject("/BaseFont"): NameObject("/STSong-Light"),
                            NameObject("/Encoding"): NameObject("/UniGB-UCS2-H"),
                            NameObject("/DescendantFonts"): ArrayObject([cid])})
    page[NameObject("/Resources")] = DictionaryObject({NameObject("/Font"): DictionaryObject({NameObject("/F1"): font, NameObject("/F2"): cjk})})
    content = DecodedStreamObject()
    chinese = "中文课程转换".encode("utf-16-be").hex().encode("ascii")
    content.set_data(b"BT /F1 16 Tf 50 700 Td (MarkItDown PDF conversion) Tj ET\n"
                     b"BT /F2 16 Tf 50 650 Td <" + chinese + b"> Tj ET")
    page[NameObject("/Contents")] = writer._add_object(content)
    writer.write(path)


class MarkdownTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def test_recursive_conversion_keeps_sources_and_isolates_failure(self):
        nested = self.folder / "中文 子目录"
        nested.mkdir()
        good = nested / "课程.pdf"
        good.write_bytes(b"fixture")
        (self.folder / "broken.pdf").write_bytes(b"bad")
        (self.folder / "ignore.docx").write_bytes(b"ignore")
        def convert(source, staged, timeout, cancel):
            if Path(source).name == "broken.pdf":
                raise ValueError("damaged PDF")
            Path(staged).write_text("# 中文课程\n", encoding="utf-8")
        with patch("markdown_backend._convert_pdf", side_effect=convert):
            report = convert_pdf_folder_report(str(self.folder), recursive=True)
        self.assertEqual((report.total, report.succeeded, report.failed), (2, 1, 1))
        self.assertEqual(good.read_bytes(), b"fixture")
        self.assertEqual(good.with_suffix(".md").read_text(encoding="utf-8"), "# 中文课程\n")
        self.assertIn("PDF", report.summary())
        self.assertFalse(list(self.folder.rglob(".md-convert-*")))

    def test_empty_output_does_not_replace_existing_markdown(self):
        source = self.folder / "blank.pdf"
        source.write_bytes(b"fixture")
        target = source.with_suffix(".md")
        target.write_text("existing", encoding="utf-8")
        with patch("markdown_backend._convert_pdf"):
            report = convert_pdf_folder_report(str(self.folder))
        self.assertEqual(report.failed, 1)
        self.assertEqual(target.read_text(encoding="utf-8"), "existing")
        self.assertFalse(list(self.folder.glob(".md-convert-*")))

    def test_cancel_before_start_and_invalid_timeout(self):
        report = convert_pdf_folder_report(str(self.folder), cancel_callback=lambda: True)
        self.assertTrue(report.cancelled)
        for timeout in (0, float("nan"), float("inf")):
            with self.assertRaises(ValueError):
                convert_pdf_folder_report(str(self.folder), timeout=timeout)

    def test_retry_rejects_outside_output(self):
        source = self.folder / "test.pdf"
        source.write_bytes(b"fixture")
        with self.assertRaises(ValueError):
            convert_pdf_folder_report(str(self.folder), only_files=[str(source)],
                                      output_paths={str(source): str(self.folder.parent / "outside.md")})

    def test_timeout_and_active_cancellation_reap_workers(self):
        initial = {p.pid for p in multiprocessing.active_children()}
        with patch("markdown_backend._markdown_worker", stalled_markdown_worker):
            with self.assertRaises(TimeoutError):
                _convert_pdf("unused.pdf", "unused.md", 0.2, None)
            start = time.monotonic()
            with self.assertRaises(ConversionCancelled):
                _convert_pdf("unused.pdf", "unused.md", 10, lambda: time.monotonic() - start > 0.2)
        self.assertEqual({p.pid for p in multiprocessing.active_children()}, initial)

    def test_real_markitdown_text_blank_damaged_and_encrypted(self):
        from pypdf import PdfWriter
        source = self.folder / "中文 课程.pdf"
        create_text_pdf(source)
        original = source.read_bytes()
        blank = PdfWriter()
        blank.add_blank_page(600, 800)
        blank.write(self.folder / "blank.pdf")
        (self.folder / "broken.pdf").write_bytes(b"not a pdf")
        encrypted = PdfWriter(clone_from=source)
        encrypted.encrypt("secret-password")
        encrypted.write(self.folder / "encrypted.pdf")
        report = convert_pdf_folder_report(str(self.folder), log_callback=lambda _: None)
        self.assertEqual((report.succeeded, report.failed), (1, 3))
        self.assertIn("MarkItDown PDF conversion", source.with_suffix(".md").read_text(encoding="utf-8"))
        self.assertIn("中文课程转换", source.with_suffix(".md").read_text(encoding="utf-8"))
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(len(list(self.folder.glob("*.md"))), 1)
        self.assertFalse(list(self.folder.glob(".md-convert-*")))


if __name__ == "__main__":
    unittest.main()
