import os
import tempfile
import unittest
from types import SimpleNamespace
from pathlib import Path
from unittest.mock import patch

import run


def pdf_bytes(text="fixture"):
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>",
    ]
    data = b"%PDF-1.4\n% " + text.encode() + b"\n"
    offsets = [0]
    for number, obj in enumerate(objects, 1):
        offsets.append(len(data))
        data += f"{number} 0 obj\n".encode() + obj + b"\nendobj\n"
    start = len(data)
    data += b"xref\n0 4\n0000000000 65535 f \n"
    for offset in offsets[1:]:
        data += f"{offset:010d} 00000 n \n".encode()
    return data + f"trailer\n<< /Size 4 /Root 1 0 R >>\nstartxref\n{start}\n%%EOF\n".encode()


class FakeWorker:
    exporter = None

    def __init__(self, *args, **kwargs):
        pass

    def convert(self, source, target, cancel_callback=None):
        if self.exporter:
            self.exporter(source, target)
        else:
            Path(target).write_bytes(pdf_bytes(Path(source).suffix))

    def close(self):
        pass


class ConversionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.worker_patch = patch.object(run, "OfficeWorker", FakeWorker, create=True)
        self.worker_patch.start()
        self.addCleanup(self.worker_patch.stop)
        FakeWorker.exporter = None

    def source(self, name="测试.docx"):
        path = self.folder / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"source content")
        return path

    def report(self, **kwargs):
        return run.convert_folder_report(str(self.folder), log_callback=lambda _: None, **kwargs)

    def test_invalid_export_is_failure_and_preserves_source_and_existing_pdf(self):
        source = self.source()
        target = source.with_suffix(".pdf")
        previous = pdf_bytes("previous")
        target.write_bytes(previous)
        FakeWorker.exporter = staticmethod(lambda src, dst: Path(dst).write_bytes(b"broken PDF"))
        report = self.report(delete_source=True)
        self.assertEqual(report.failed, 1)
        self.assertEqual(report.succeeded, 0)
        self.assertTrue(source.exists())
        self.assertEqual(target.read_bytes(), previous)
        self.assertEqual(sorted(p.name for p in self.folder.iterdir()), ["测试.docx", "测试.pdf"])

    def test_legacy_batch_does_not_count_or_delete_invalid_export(self):
        source = self.source()
        document = SimpleNamespace(
            SaveAs=lambda path, fmt: Path(path).write_bytes(b"broken PDF"),
            Close=lambda *args: None,
        )
        office = SimpleNamespace(Documents=SimpleNamespace(Open=lambda *a, **kw: document), Quit=lambda: None)
        FakeWorker.exporter = staticmethod(lambda src, dst: Path(dst).write_bytes(b"broken PDF"))
        with patch.object(run, "_get_word", return_value=office, create=True):
            count = run.batch_doc_to_pdf(str(self.folder), delete_source=True, log_callback=lambda _: None)
        self.assertEqual(count, 0)
        self.assertTrue(source.exists())

    def test_default_conversion_preserves_source(self):
        source = self.source()
        report = self.report()
        self.assertEqual(report.succeeded, 1)
        self.assertTrue(source.exists())
        self.assertTrue(source.with_suffix(".pdf").exists())

    def test_changed_source_is_preserved_even_when_deletion_is_enabled(self):
        source = self.source()
        def export(src, dst):
            Path(dst).write_bytes(pdf_bytes())
            Path(src).write_bytes(b"edited while Office was exporting")
        FakeWorker.exporter = staticmethod(export)
        report = self.report(delete_source=True)
        self.assertEqual(report.succeeded, 1)
        self.assertTrue(source.exists())
        self.assertTrue(report.results[0].warning)

    def test_valid_export_can_delete_unchanged_source_when_opted_in(self):
        source = self.source()
        report = self.report(delete_source=True)
        self.assertEqual(report.succeeded, 1)
        self.assertFalse(source.exists())
        self.assertTrue(source.with_suffix(".pdf").exists())

    def test_same_stem_formats_never_overwrite_each_other(self):
        self.source("report.pptx")
        self.source("report.docx")
        report = self.report()
        outputs = [Path(item.output) for item in report.results]
        self.assertEqual(report.succeeded, 2)
        self.assertEqual(len(set(outputs)), 2)
        self.assertEqual({p.name for p in outputs}, {"report.pptx.pdf", "report.docx.pdf"})
        self.assertNotEqual(outputs[0].read_bytes(), outputs[1].read_bytes())

    def test_collision_with_natural_double_extension_is_also_disambiguated(self):
        for name in ("a.docx", "a.pptx", "a.docx.docx"):
            self.source(name)
        report = self.report()
        self.assertEqual(report.succeeded, 3)
        self.assertEqual(len({item.output.lower() for item in report.results}), 3)

    def test_retry_preserves_pdfs_after_colliding_successful_sources_are_deleted(self):
        for name in ("a.docx", "a.pptx", "a.docx.docx"):
            self.source(name)
        def export(src, dst):
            if Path(src).name == "a.docx.docx":
                raise RuntimeError("damaged")
            Path(dst).write_bytes(pdf_bytes(Path(src).name))
        FakeWorker.exporter = staticmethod(export)
        first = self.report(delete_source=True)
        saved = {item.output: Path(item.output).read_bytes() for item in first.results if item.status == "success"}
        self.assertEqual(len(saved), 2)
        FakeWorker.exporter = None
        second = self.report(delete_source=True, only_files=first.failed_files)
        self.assertEqual(second.succeeded, 1)
        for path, content in saved.items():
            self.assertEqual(Path(path).read_bytes(), content)
        self.assertNotIn(second.results[0].output, saved)

    def test_retry_can_reuse_original_output_mapping(self):
        source = self.source("a.docx")
        target = self.folder / "original-choice.pdf"
        target.write_bytes(pdf_bytes("old version"))
        report = self.report(only_files=[str(source)], output_paths={str(source): str(target)})
        self.assertEqual(report.results[0].output, str(target))
        self.assertEqual(target.read_bytes(), pdf_bytes(".docx"))

    def test_mixed_batch_progress_is_monotonic_and_includes_failures(self):
        self.source("a.pptx")
        self.source("b.docx")
        self.source("c.docx")
        def export(src, dst):
            if Path(src).name == "b.docx":
                raise OSError("damaged document")
            Path(dst).write_bytes(pdf_bytes())
        FakeWorker.exporter = staticmethod(export)
        progress = []
        report = self.report(progress_callback=lambda cur, total, name: progress.append((cur, total)))
        self.assertEqual((report.succeeded, report.failed), (2, 1))
        self.assertEqual(progress, [(0, 3), (1, 3), (2, 3), (3, 3)])

    def test_retry_recovers_transient_export_failure(self):
        self.source()
        attempts = []
        def export(src, dst):
            attempts.append(src)
            if len(attempts) == 1:
                raise OSError("temporary disconnection")
            Path(dst).write_bytes(pdf_bytes())
        FakeWorker.exporter = staticmethod(export)
        report = self.report()
        self.assertEqual(report.succeeded, 1)
        self.assertEqual(len(attempts), 2)

    def test_cancellation_stops_before_next_file(self):
        self.source("a.docx")
        self.source("b.docx")
        finished = []
        def export(src, dst):
            Path(dst).write_bytes(pdf_bytes())
            finished.append(src)
        FakeWorker.exporter = staticmethod(export)
        report = self.report(cancel_callback=lambda: bool(finished))
        self.assertEqual(report.succeeded, 1)
        self.assertTrue(report.cancelled)
        self.assertEqual(report.failed, 0)
        self.assertEqual(report.pending, 1)

    def test_only_failed_files_are_retried(self):
        good = self.source("good.docx")
        bad = self.source("bad.docx")
        FakeWorker.exporter = staticmethod(lambda src, dst: (_ for _ in ()).throw(OSError("failure")))
        first = self.report()
        FakeWorker.exporter = None
        bad.unlink()
        second = self.report(only_files=first.failed_files)
        self.assertEqual(second.succeeded, 1)
        self.assertEqual(second.failed, 1)
        self.assertTrue(good.exists())

    def test_skips_office_lock_files_and_directories_with_office_extension(self):
        self.source("~$locked.docx")
        (self.folder / "not-a-file.docx").mkdir()
        self.source("real.docx")
        report = self.report()
        self.assertEqual(report.total, 1)

    def test_empty_folder_and_invalid_folder(self):
        report = self.report()
        self.assertEqual(report.total, 0)
        with self.assertRaises(FileNotFoundError):
            run.convert_folder_report(str(self.folder / "missing"))

    def test_missing_backend_is_reported_as_failure(self):
        self.source()
        FakeWorker.exporter = staticmethod(lambda src, dst: (_ for _ in ()).throw(RuntimeError("Office unavailable")))
        report = self.report()
        self.assertEqual(report.failed, 1)
        self.assertIn("Office unavailable", report.results[0].error)

    def test_legacy_tuple_api_and_recursive_folders(self):
        source = self.source("子目录/演示.pptx")
        result = run.convert_folder(str(self.folder), recursive=True, log_callback=lambda _: None)
        self.assertEqual(result, (1, 0))
        self.assertTrue(source.exists())
        self.assertTrue(source.with_suffix(".pdf").exists())


if __name__ == "__main__":
    unittest.main()
