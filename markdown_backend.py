"""Local PDF to Markdown conversion using Microsoft's MarkItDown."""

import math
import multiprocessing
import os
import tempfile
import time

from run import ConversionReport, FileResult, _iter_files, _log
from office_backend import ConversionCancelled


class MarkdownReport(ConversionReport):
    def summary(self):
        if self.total == 0:
            return "已取消，尚未开始转换。" if self.cancelled else "未找到 PDF 文件。"
        prefix = "已取消" if self.cancelled else "处理结束"
        return (f"{prefix}：PDF → Markdown 成功 {self.succeeded} 个，"
                f"失败 {self.failed} 个，未处理 {self.pending} 个。")


def _markdown_worker(connection, source, staged):
    try:
        from pypdf import PdfReader

        with open(source, "rb") as stream:
            if stream.read(5) != b"%PDF-":
                raise ValueError("文件头不是 PDF，无法转换。")
            reader = PdfReader(stream, strict=True)
            if reader.is_encrypted:
                raise ValueError("PDF 已加密，请先解密后转换。")
            if not reader.pages:
                raise ValueError("PDF 没有页面。")
        from markitdown import MarkItDown

        result = MarkItDown(enable_plugins=False).convert(source)
        content = result.text_content
        if not content or not content.strip():
            raise ValueError("未提取到文字；扫描件或图片型 PDF 需要 OCR，当前模式不支持。")
        with open(staged, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(content.rstrip() + "\n")
        connection.send(None)
    except Exception as exc:
        try:
            connection.send(f"{type(exc).__name__}: {exc}")
        except (BrokenPipeError, EOFError, OSError):
            pass
    finally:
        connection.close()


def _convert_pdf(source, staged, timeout, cancel_callback):
    context = multiprocessing.get_context("spawn")
    receiver, sender = context.Pipe(duplex=False)
    process = context.Process(target=_markdown_worker, args=(sender, source, staged))
    started = False
    try:
        process.start()
        started = True
        sender.close()
        deadline = time.monotonic() + timeout
        while True:
            if cancel_callback and cancel_callback():
                raise ConversionCancelled()
            if time.monotonic() >= deadline:
                raise TimeoutError(f"PDF 转 Markdown 超过 {timeout:g} 秒")
            if receiver.poll(0.05):
                try:
                    error = receiver.recv()
                except EOFError as exc:
                    raise RuntimeError("Markdown 转换进程意外退出") from exc
                if error is not None:
                    raise RuntimeError(error)
                return
            if not process.is_alive():
                raise RuntimeError(f"Markdown 转换进程意外退出（{process.exitcode}）")
    finally:
        receiver.close()
        sender.close()
        if started:
            process.join(0.1)
            if process.is_alive():
                process.terminate()
                process.join(2)
            if process.is_alive():
                process.kill()
                process.join()
            process.close()


def convert_pdf_folder_report(folder, recursive=False, log_callback=None,
                              cancel_callback=None, progress_callback=None,
                              timeout=120, only_files=None, output_paths=None):
    folder = os.path.abspath(os.path.normpath(folder))
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"文件夹不存在：{folder}")
    if not math.isfinite(timeout) or timeout <= 0:
        raise ValueError("单文件超时时间必须是大于 0 的有限秒数。")
    if cancel_callback and cancel_callback():
        return MarkdownReport(0, cancelled=True)
    if only_files is None:
        files = [os.path.join(folder, name) for name in _iter_files(folder, (".pdf",), recursive)]
    else:
        files = sorted({os.path.abspath(os.path.join(folder, name)) for name in only_files}, key=str.casefold)
    targets = {}
    for source in files:
        if os.path.normcase(os.path.commonpath([folder, source])) != os.path.normcase(folder) or not source.lower().endswith(".pdf"):
            raise ValueError(f"源文件必须是所选文件夹中的 PDF：{source}")
        target = os.path.splitext(source)[0] + ".md"
        if only_files is not None and output_paths and source in output_paths:
            target = os.path.abspath(output_paths[source])
        if os.path.normcase(os.path.commonpath([folder, target])) != os.path.normcase(folder) or not target.lower().endswith(".md"):
            raise ValueError(f"输出必须是所选文件夹中的 Markdown：{target}")
        targets[source] = target
    if len({os.path.normcase(target) for target in targets.values()}) != len(files):
        raise ValueError("Markdown 输出路径重复。")
    report = MarkdownReport(len(files))
    if progress_callback:
        progress_callback(0, report.total, "")
    for source in files:
        if cancel_callback and cancel_callback():
            report.cancelled = True
            break
        target = targets[source]
        relative = os.path.relpath(source, folder)
        result = FileResult(source, target, "pdf", "failed")
        _log(f"[{len(report.results) + 1}/{report.total}] PDF → Markdown：{relative}", log_callback)
        try:
            for attempt in range(2):
                staged = None
                try:
                    if not os.path.isfile(source):
                        raise FileNotFoundError(f"源 PDF 不存在：{source}")
                    descriptor, staged = tempfile.mkstemp(prefix=".md-convert-", suffix=".md", dir=os.path.dirname(target))
                    os.close(descriptor)
                    _convert_pdf(source, staged, timeout, cancel_callback)
                    if cancel_callback and cancel_callback():
                        raise ConversionCancelled()
                    with open(staged, encoding="utf-8") as stream:
                        if not stream.read().strip():
                            raise ValueError("未提取到文字；扫描件需要 OCR，当前模式不支持。")
                    os.replace(staged, target)
                    result.status = "success"
                    break
                except ConversionCancelled:
                    raise
                except Exception:
                    if attempt == 1:
                        raise
                    _log(f"首次转换失败，重试：{relative}", log_callback)
                finally:
                    if staged and os.path.exists(staged):
                        os.remove(staged)
            _log(f"完成：{os.path.relpath(target, folder)}", log_callback)
        except ConversionCancelled:
            report.cancelled = True
            break
        except Exception as exc:
            result.error = str(exc)
            _log(f"失败：{relative}；{exc}", log_callback)
        report.results.append(result)
        if progress_callback:
            progress_callback(len(report.results), report.total, relative)
    return report
