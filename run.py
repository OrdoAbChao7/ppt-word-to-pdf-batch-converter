import math
import multiprocessing
import os
import shutil
import stat
import sys
import tempfile
from collections import Counter
from dataclasses import dataclass, field
from typing import Callable, Iterable

from pypdf import PdfReader

from office_backend import ConversionCancelled, OfficeUnavailable, OfficeWorker


LogCallback = Callable[[str], None]
CancelCallback = Callable[[], bool]
ProgressCallback = Callable[[int, int, str], None]

PPT_EXTENSIONS = (".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm")
DOC_EXTENSIONS = (".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf")


@dataclass
class FileResult:
    source: str
    output: str
    kind: str
    status: str
    error: str = ""
    warning: str = ""


@dataclass
class ConversionReport:
    total: int
    results: list[FileResult] = field(default_factory=list)
    cancelled: bool = False

    @property
    def succeeded(self):
        return sum(item.status == "success" for item in self.results)

    @property
    def failed(self):
        return sum(item.status == "failed" for item in self.results)

    @property
    def pending(self):
        return self.total - self.succeeded - self.failed

    @property
    def failed_files(self):
        return [item.source for item in self.results if item.status == "failed"]

    @property
    def ppt_count(self):
        return sum(item.kind == "ppt" and item.status == "success" for item in self.results)

    @property
    def doc_count(self):
        return sum(item.kind == "doc" and item.status == "success" for item in self.results)

    def summary(self):
        prefix = "已取消" if self.cancelled else "处理结束"
        if self.total == 0:
            return "已取消，尚未开始转换。" if self.cancelled else "未找到支持的 PPT / Word 文件。"
        return (f"{prefix}：成功 {self.succeeded} 个（PPT {self.ppt_count}，Word {self.doc_count}），"
                f"失败 {self.failed} 个，未处理 {self.pending} 个。")


def _log(message: str, callback: LogCallback | None = None) -> None:
    if callback:
        callback(message)
    else:
        print(message)


def _iter_files(folder: str, extensions: Iterable[str], recursive: bool = False) -> list[str]:
    ext_tuple = tuple(e.lower() for e in extensions)
    matched = []
    if recursive:
        def on_error(error):
            raise error
        entries = os.walk(folder, onerror=on_error)
    else:
        entries = [(folder, [], os.listdir(folder))]
    for root, _dirs, files in entries:
        for name in files:
            path = os.path.join(root, name)
            if name.lower().endswith(ext_tuple) and not name.startswith("~") and os.path.isfile(path):
                matched.append(os.path.relpath(path, folder))
    return sorted(matched, key=str.casefold)


def _is_file_writable(filepath: str) -> bool:
    if not os.path.exists(filepath):
        return True
    try:
        with open(filepath, "r+b"):
            return True
    except OSError:
        return False


def _safe_remove_file(filepath: str) -> None:
    try:
        os.remove(filepath)
    except PermissionError:
        os.chmod(filepath, stat.S_IWRITE)
        os.remove(filepath)


def _output_paths(files):
    """Reserve outputs across BOTH Office formats, including double extensions."""
    stems = Counter(os.path.normcase(os.path.splitext(path)[0]) for path in files)
    desired = {
        path: (path if stems[os.path.normcase(os.path.splitext(path)[0])] > 1
               else os.path.splitext(path)[0]) + ".pdf"
        for path in files
    }
    names = Counter(os.path.normcase(path) for path in desired.values())
    used = set()
    outputs = {}
    for source, target in desired.items():
        base = os.path.splitext(target)[0]
        candidate = target
        counter = 2
        while os.path.normcase(candidate) in used or (
            candidate != target and os.path.normcase(candidate) in names
        ):
            candidate = f"{base}_{counter}.pdf"
            counter += 1
        outputs[source] = candidate
        used.add(os.path.normcase(candidate))
    return outputs


def _validate_pdf(path):
    try:
        with open(path, "rb") as stream:
            if stream.read(5) != b"%PDF-":
                raise ValueError("文件头不是 PDF")
            reader = PdfReader(stream, strict=True)
            if reader.is_encrypted or len(reader.pages) == 0:
                raise ValueError("PDF 加密或没有页面")
            for page in reader.pages:
                _ = page.mediabox
    except Exception as exc:
        raise ValueError(f"导出的 PDF 校验失败，源文件已保留：{exc}") from exc


def _convert_files(folder, files, outputs, delete_source, timeout, log_callback,
                   cancel_callback, progress_callback):
    report = ConversionReport(total=len(files))
    workers = {}
    unavailable = {}
    if progress_callback:
        progress_callback(0, report.total, "")
    try:
        for source in files:
            if cancel_callback and cancel_callback():
                report.cancelled = True
                break
            kind = "ppt" if source.lower().endswith(PPT_EXTENSIONS) else "doc"
            target = outputs[source]
            relative = os.path.relpath(source, folder)
            _log(f"[{report.succeeded + report.failed + 1}/{report.total}] 正在转换：{relative}", log_callback)
            result = FileResult(source, target, kind, "failed")
            staged = None
            try:
                if not os.path.isfile(source):
                    raise FileNotFoundError(f"源文件不存在或已移动：{source}")
                source_stat = os.stat(source)
                source_version = (source_stat.st_ino, source_stat.st_size, source_stat.st_mtime_ns)
                if not _is_file_writable(target):
                    raise PermissionError(f"目标 PDF 被占用或只读，请关闭阅读器或检查权限：{target}")
                if kind in unavailable:
                    raise OfficeUnavailable(unavailable[kind])
                for attempt in range(2):
                    fd, staged = tempfile.mkstemp(prefix=".pdf-convert-", suffix=".pdf", dir=os.path.dirname(target))
                    os.close(fd)
                    try:
                        if kind not in workers:
                            workers[kind] = OfficeWorker(kind, timeout=timeout)
                        workers[kind].convert(source, staged, cancel_callback)
                        _validate_pdf(staged)
                        os.replace(staged, target)
                        staged = None
                        result.status = "success"
                        break
                    except ConversionCancelled:
                        raise
                    except Exception as exc:
                        worker = workers.pop(kind, None)
                        if worker:
                            worker.close()
                        if isinstance(exc, OfficeUnavailable):
                            unavailable[kind] = str(exc)
                            raise
                        if attempt == 1:
                            raise
                        _log(f"首次转换失败，将重新连接 Office 后重试：{relative}；{exc}", log_callback)
                    finally:
                        if staged and os.path.exists(staged):
                            os.remove(staged)
                        staged = None
                _log(f"完成：{os.path.relpath(target, folder)}", log_callback)
                if delete_source:
                    try:
                        current_stat = os.stat(source)
                        if (current_stat.st_ino, current_stat.st_size, current_stat.st_mtime_ns) != source_version:
                            raise OSError("源文件在转换期间发生变化，已保留新版本")
                        _safe_remove_file(source)
                        _log(f"已删除源文件：{relative}", log_callback)
                    except OSError as exc:
                        result.warning = f"PDF 已生成，源文件删除失败：{exc}"
                        _log(result.warning, log_callback)
            except ConversionCancelled:
                report.cancelled = True
                break
            except Exception as exc:
                result.error = str(exc)
                _log(f"失败：{relative}；{exc}", log_callback)
            report.results.append(result)
            if progress_callback:
                progress_callback(report.succeeded + report.failed, report.total, relative)
    finally:
        for worker in workers.values():
            worker.close()
    _log(report.summary(), log_callback)
    return report


def convert_folder_report(
    folder: str, recursive: bool = False, flatten: bool = False,
    delete_source: bool = False, log_callback: LogCallback | None = None,
    cancel_callback: CancelCallback | None = None,
    progress_callback: ProgressCallback | None = None,
    timeout: float = 120, only_files: Iterable[str] | None = None,
    output_paths: dict[str, str] | None = None,
) -> ConversionReport:
    folder = os.path.abspath(os.path.normpath(folder))
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"文件夹不存在：{folder}")
    if not math.isfinite(timeout) or timeout <= 0:
        raise ValueError("单文件超时时间必须是大于 0 的有限秒数。")
    if cancel_callback and cancel_callback():
        return ConversionReport(0, cancelled=True)
    if flatten and only_files is None:
        flatten_folders(folder, log_callback, cancel_callback)
    files = [os.path.join(folder, path) for path in _iter_files(folder, PPT_EXTENSIONS + DOC_EXTENSIONS, recursive)]
    if only_files is not None:
        selected = sorted({os.path.abspath(os.path.join(folder, path)) for path in only_files}, key=str.casefold)
        for path in selected:
            if os.path.commonpath([folder, path]) != folder or not path.lower().endswith(PPT_EXTENSIONS + DOC_EXTENSIONS):
                raise ValueError(f"重试文件不在所选文件夹中或格式不受支持：{path}")
        all_files = sorted(set(files) | set(selected), key=str.casefold)
        files = selected
    else:
        all_files = files
    outputs = _output_paths(all_files)
    if only_files is not None:
        reserved = {os.path.normcase(path) for path in outputs.values()}
        for source in files:
            if output_paths and source in output_paths:
                target = os.path.abspath(output_paths[source])
                if os.path.normcase(os.path.commonpath([folder, target])) != os.path.normcase(folder) or not target.lower().endswith(".pdf"):
                    raise ValueError(f"重试输出路径必须是所选文件夹中的 PDF：{target}")
            else:
                target = outputs[source]
                base = os.path.splitext(target)[0]
                counter = 2
                while os.path.exists(target):
                    target = f"{base}_{counter}.pdf"
                    counter += 1
                    while os.path.normcase(target) in reserved:
                        target = f"{base}_{counter}.pdf"
                        counter += 1
            outputs[source] = target
            reserved.add(os.path.normcase(target))
        if len({os.path.normcase(outputs[path]) for path in files}) != len(files):
            raise ValueError("重试输出路径重复，请重新开始转换。")
    return _convert_files(folder, files, outputs, delete_source, timeout,
                          log_callback, cancel_callback, progress_callback)


def convert_folder(folder: str, recursive: bool = False, flatten: bool = False,
                   delete_source: bool = False, log_callback: LogCallback | None = None,
                   cancel_callback: CancelCallback | None = None,
                   progress_callback: ProgressCallback | None = None) -> tuple[int, int]:
    report = convert_folder_report(folder, recursive, flatten, delete_source, log_callback,
                                   cancel_callback, progress_callback)
    return report.ppt_count, report.doc_count


def _legacy_batch(folder, extensions, recursive, delete_source, log_callback, cancel_callback, progress_callback):
    folder = os.path.abspath(folder)
    all_files = [os.path.join(folder, path) for path in _iter_files(folder, PPT_EXTENSIONS + DOC_EXTENSIONS, recursive)]
    files = [path for path in all_files if path.lower().endswith(extensions)]
    report = _convert_files(folder, files, _output_paths(all_files), delete_source, 120,
                            log_callback, cancel_callback, progress_callback)
    return report.succeeded


def batch_ppt_to_pdf(folder: str, recursive: bool = False, delete_source: bool = False,
                     log_callback=None, cancel_callback=None, progress_callback=None) -> int:
    return _legacy_batch(folder, PPT_EXTENSIONS, recursive, delete_source, log_callback, cancel_callback, progress_callback)


def batch_doc_to_pdf(folder: str, recursive: bool = False, delete_source: bool = False,
                     log_callback=None, cancel_callback=None, progress_callback=None) -> int:
    return _legacy_batch(folder, DOC_EXTENSIONS, recursive, delete_source, log_callback, cancel_callback, progress_callback)


def flatten_folders(folder: str, log_callback: LogCallback | None = None, cancel_callback: CancelCallback | None = None) -> tuple[int, int]:
    """Move Office files from subfolders into the selected folder and remove empty folders."""
    folder = os.path.abspath(os.path.normpath(folder))
    _log(f"开始整理子文件夹中的 Office 文件：{folder}", log_callback)

    target_extensions = PPT_EXTENSIONS + DOC_EXTENSIONS
    files_to_move: list[str] = []
    for root, _dirs, files in os.walk(folder, topdown=False):
        if root == folder:
            continue

        for file in files:
            if file.lower().endswith(target_extensions) and not file.startswith("~"):
                files_to_move.append(os.path.join(root, file))

    moved_count = 0
    for file_path in files_to_move:
        if cancel_callback and cancel_callback():
            break
        try:
            file_name = os.path.basename(file_path)
            new_path = os.path.join(folder, file_name)
            counter = 1
            while os.path.exists(new_path):
                name, ext = os.path.splitext(file_name)
                new_path = os.path.join(folder, f"{name}_{counter}{ext}")
                counter += 1

            shutil.move(file_path, new_path)
            moved_count += 1
            _log(f"已移动：{file_name}", log_callback)
        except Exception as exc:
            _log(f"移动失败：{file_path}；原因：{exc}", log_callback)

    deleted_count = 0
    for root, _dirs, _files in os.walk(folder, topdown=False):
        if root == folder:
            continue

        try:
            os.rmdir(root)
            deleted_count += 1
            _log(f"已删除空文件夹：{root}", log_callback)
        except OSError:
            pass

    _log(
        f"子文件夹整理完成，共移动 {moved_count} 个 Office 文件，删除 {deleted_count} 个空文件夹。",
        log_callback,
    )
    return moved_count, deleted_count



def main() -> int:
    folder = input("请输入需要批量转换的文件夹路径：").strip().strip('"')
    recursive = input("是否包含子文件夹？(y/n)：").lower() == "y"
    flatten = False if recursive else input("是否集中移动子文件夹中的文档？(y/n)：").lower() == "y"
    delete_source = input("转换成功后是否删除源文件？(y/n)：").lower() == "y"
    try:
        report = convert_folder_report(folder, recursive, flatten, delete_source)
        return 1 if report.failed or report.cancelled else 0
    except (OSError, ValueError) as exc:
        print(f"处理失败：{exc}")
        return 1


if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.exit(main())
