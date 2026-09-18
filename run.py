import collections
import os
import shutil
import stat
import sys
import time
from typing import Callable, Iterable

import pythoncom
import win32com.client


LogCallback = Callable[[str], None]
CancelCallback = Callable[[], bool]
ProgressCallback = Callable[[int, int, str], None]

PPT_EXTENSIONS = (".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm")
DOC_EXTENSIONS = (".doc", ".docx", ".docm", ".dot", ".dotx", ".rtf")


def _log(message: str, callback: LogCallback | None = None) -> None:
    if callback:
        callback(message)
    else:
        print(message)


def _iter_files(folder: str, extensions: Iterable[str], recursive: bool = False) -> list[str]:
    ext_tuple = tuple(e.lower() for e in extensions)
    if not recursive:
        return [
            file
            for file in os.listdir(folder)
            if file.lower().endswith(ext_tuple) and not file.startswith("~")
        ]

    matched: list[str] = []
    for root, _dirs, files in os.walk(folder):
        for file in files:
            if file.lower().endswith(ext_tuple) and not file.startswith("~"):
                rel_path = os.path.relpath(os.path.join(root, file), folder)
                matched.append(rel_path)
    return matched


def _is_file_writable(filepath: str) -> bool:
    """Check whether a file can be opened for writing (not locked by another program)."""
    if not os.path.exists(filepath):
        return True
    try:
        with open(filepath, "r+b"):
            return True
    except (PermissionError, OSError):
        return False


def _safe_remove_file(filepath: str) -> None:
    """Safely remove a file, clearing read-only attributes on Windows if necessary."""
    try:
        os.remove(filepath)
    except PermissionError:
        try:
            os.chmod(filepath, stat.S_IWRITE)
            os.remove(filepath)
        except Exception:
            raise


def _get_powerpoint():
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
        powerpoint.DisplayAlerts = 1  # ppAlertsNone
        return powerpoint
    except Exception:
        return None


def _get_word():
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = 0
        word.DisplayAlerts = 0  # wdAlertsNone
        word.Options.SaveNormalPrompt = False
        word.Options.SavePropertiesPrompt = False
        return word
    except Exception:
        return None


def batch_ppt_to_pdf(
    folder: str,
    recursive: bool = False,
    delete_source: bool = True,
    log_callback: LogCallback | None = None,
    cancel_callback: CancelCallback | None = None,
    progress_callback: ProgressCallback | None = None,
) -> int:
    """Convert PowerPoint files in a folder to PDF."""
    folder = os.path.abspath(os.path.normpath(folder))
    pythoncom.CoInitialize()
    ppt_files = _iter_files(folder, PPT_EXTENSIONS, recursive=recursive)
    if not ppt_files:
        _log("未找到 PPT 文件。", log_callback)
        pythoncom.CoUninitialize()
        return 0

    powerpoint = _get_powerpoint()
    if powerpoint is None:
        _log("无法启动 Microsoft PowerPoint，请确保本机已安装并正确注册 Office。", log_callback)
        pythoncom.CoUninitialize()
        return 0

    converted_files: list[str] = []
    failed_attempts: dict[str, int] = collections.defaultdict(int)
    max_retries = 2
    total_files = len(ppt_files)

    _log(f"找到 {total_files} 个 PPT 相关文件，开始转换。", log_callback)

    try:
        for retry_count in range(max_retries):
            if cancel_callback and cancel_callback():
                _log("转换已被用户取消。", log_callback)
                break

            remaining_files = [
                f for f in ppt_files
                if f not in converted_files and failed_attempts[f] < 2
            ]
            if not remaining_files:
                break

            if retry_count > 0:
                _log(f"第 {retry_count} 次重试，继续转换剩余文件（共 {len(remaining_files)} 个）。", log_callback)

            for file in remaining_files:
                if cancel_callback and cancel_callback():
                    _log("转换已被用户取消。", log_callback)
                    break

                full_path = os.path.abspath(os.path.normpath(os.path.join(folder, file)))
                pdf_path = os.path.abspath(os.path.normpath(os.path.join(folder, f"{os.path.splitext(file)[0]}.pdf")))

                # Ensure target subfolder exists if recursive
                os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

                current_index = len(converted_files) + 1
                if progress_callback:
                    progress_callback(current_index, total_files, file)

                _log(f"[{current_index}/{total_files}] 正在转换 PPT：{file}", log_callback)

                if os.path.exists(pdf_path) and not _is_file_writable(pdf_path):
                    _log(f"目标 PDF 已被其他程序占用，请先关闭：{os.path.basename(pdf_path)}", log_callback)
                    failed_attempts[file] += 1
                    continue

                presentation = None
                try:
                    if powerpoint is None:
                        powerpoint = _get_powerpoint()
                        if powerpoint is None:
                            _log("无法连接到 PowerPoint 进程，终止后续转换。", log_callback)
                            break

                    presentation = powerpoint.Presentations.Open(full_path, True, False, False)
                    presentation.SaveAs(pdf_path, 32)
                    presentation.Close()
                    presentation = None

                    converted_files.append(file)
                    _log(f"完成：{os.path.basename(pdf_path)}", log_callback)

                    if delete_source:
                        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                            try:
                                _safe_remove_file(full_path)
                                _log(f"已删除源文件：{file}", log_callback)
                            except Exception as del_err:
                                _log(f"已生成 PDF，但删除源文件失败（可能权限受限）：{file}；原因：{del_err}", log_callback)
                        else:
                            _log(f"警告：未检测到有效的生成文件，已保留源文件以防丢失：{file}", log_callback)

                except Exception as exc:
                    failed_attempts[file] += 1
                    _log(f"转换失败：{file}；原因：{exc}", log_callback)
                    if presentation is not None:
                        try:
                            presentation.Close()
                        except Exception:
                            pass
                        presentation = None
                    try:
                        if powerpoint is not None:
                            powerpoint.Quit()
                    except Exception:
                        pass
                    powerpoint = None
                    time.sleep(0.5)
                    powerpoint = _get_powerpoint()

    finally:
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass
            powerpoint = None
        pythoncom.CoUninitialize()

    _log(f"PPT 转换完成，共成功转换 {len(converted_files)} / {total_files} 个文件。", log_callback)
    return len(converted_files)


def batch_doc_to_pdf(
    folder: str,
    recursive: bool = False,
    delete_source: bool = True,
    log_callback: LogCallback | None = None,
    cancel_callback: CancelCallback | None = None,
    progress_callback: ProgressCallback | None = None,
) -> int:
    """Convert Word files in a folder to PDF."""
    folder = os.path.abspath(os.path.normpath(folder))
    pythoncom.CoInitialize()
    doc_files = _iter_files(folder, DOC_EXTENSIONS, recursive=recursive)
    if not doc_files:
        _log("未找到 Word 文件。", log_callback)
        pythoncom.CoUninitialize()
        return 0

    word = _get_word()
    if word is None:
        _log("无法启动 Microsoft Word，请确保本机已安装并正确注册 Office。", log_callback)
        pythoncom.CoUninitialize()
        return 0

    converted_files: list[str] = []
    failed_attempts: dict[str, int] = collections.defaultdict(int)
    max_retries = 2
    total_files = len(doc_files)

    _log(f"找到 {total_files} 个 Word 相关文件，开始转换。", log_callback)

    try:
        for retry_count in range(max_retries):
            if cancel_callback and cancel_callback():
                _log("转换已被用户取消。", log_callback)
                break

            remaining_files = [
                f for f in doc_files
                if f not in converted_files and failed_attempts[f] < 2
            ]
            if not remaining_files:
                break

            if retry_count > 0:
                _log(f"第 {retry_count} 次重试，继续转换剩余文件（共 {len(remaining_files)} 个）。", log_callback)

            for file in remaining_files:
                if cancel_callback and cancel_callback():
                    _log("转换已被用户取消。", log_callback)
                    break

                full_path = os.path.abspath(os.path.normpath(os.path.join(folder, file)))
                pdf_path = os.path.abspath(os.path.normpath(os.path.join(folder, f"{os.path.splitext(file)[0]}.pdf")))

                # Ensure target subfolder exists if recursive
                os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

                current_index = len(converted_files) + 1
                if progress_callback:
                    progress_callback(current_index, total_files, file)

                _log(f"[{current_index}/{total_files}] 正在转换 Word：{file}", log_callback)

                if os.path.exists(pdf_path) and not _is_file_writable(pdf_path):
                    _log(f"目标 PDF 已被其他程序占用，请先关闭：{os.path.basename(pdf_path)}", log_callback)
                    failed_attempts[file] += 1
                    continue

                doc = None
                try:
                    if word is None:
                        word = _get_word()
                        if word is None:
                            _log("无法连接到 Word 进程，终止后续转换。", log_callback)
                            break

                    doc = word.Documents.Open(
                        full_path,
                        ReadOnly=True,
                        ConfirmConversions=False,
                        AddToRecentFiles=False,
                    )
                    doc.SaveAs(pdf_path, 17)
                    doc.Close(0)
                    doc = None

                    converted_files.append(file)
                    _log(f"完成：{os.path.basename(pdf_path)}", log_callback)

                    if delete_source:
                        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                            try:
                                _safe_remove_file(full_path)
                                _log(f"已删除源文件：{file}", log_callback)
                            except Exception as del_err:
                                _log(f"已生成 PDF，但删除源文件失败（可能权限受限）：{file}；原因：{del_err}", log_callback)
                        else:
                            _log(f"警告：未检测到有效的生成文件，已保留源文件以防丢失：{file}", log_callback)

                except Exception as exc:
                    failed_attempts[file] += 1
                    _log(f"转换失败：{file}；原因：{exc}", log_callback)
                    if doc is not None:
                        try:
                            doc.Close(0)
                        except Exception:
                            pass
                        doc = None
                    try:
                        if word is not None:
                            word.Quit()
                    except Exception:
                        pass
                    word = None
                    time.sleep(0.5)
                    word = _get_word()

    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
            word = None
        pythoncom.CoUninitialize()

    _log(f"Word 转换完成，共成功转换 {len(converted_files)} / {total_files} 个文件。", log_callback)
    return len(converted_files)


def flatten_folders(folder: str, log_callback: LogCallback | None = None) -> tuple[int, int]:
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


def convert_folder(
    folder: str,
    recursive: bool = False,
    flatten: bool = False,
    delete_source: bool = True,
    log_callback: LogCallback | None = None,
    cancel_callback: CancelCallback | None = None,
    progress_callback: ProgressCallback | None = None,
) -> tuple[int, int]:
    folder = os.path.abspath(os.path.normpath(folder))
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"文件夹不存在：{folder}")

    if flatten:
        flatten_folders(folder, log_callback)

    if cancel_callback and cancel_callback():
        return 0, 0

    ppt_count = batch_ppt_to_pdf(
        folder,
        recursive=recursive,
        delete_source=delete_source,
        log_callback=log_callback,
        cancel_callback=cancel_callback,
        progress_callback=progress_callback,
    )

    if cancel_callback and cancel_callback():
        return ppt_count, 0

    doc_count = batch_doc_to_pdf(
        folder,
        recursive=recursive,
        delete_source=delete_source,
        log_callback=log_callback,
        cancel_callback=cancel_callback,
        progress_callback=progress_callback,
    )

    _log(f"全部完成：PPT {ppt_count} 个，Word {doc_count} 个。", log_callback)
    return ppt_count, doc_count


def main() -> int:
    folder_path = input("请输入需要批量转换的文件夹路径：").strip().strip('"')
    if not os.path.exists(folder_path):
        print(f"错误：文件夹不存在：{folder_path}")
        return 1

    recursive_choice = input("是否包含子文件夹（原地转换，保留原有目录结构）？(y/n)：").lower()
    flatten_choice = "n"
    if recursive_choice != "y":
        flatten_choice = input("是否整理子文件夹，把里面的文件集中到主文件夹？(y/n)：").lower()
    delete_choice = input("转换成功后是否删除源文件？(y/n)：").lower()

    convert_folder(
        folder_path,
        recursive=recursive_choice == "y",
        flatten=flatten_choice == "y",
        delete_source=delete_choice == "y",
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
