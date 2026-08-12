import os
import shutil
import sys
from typing import Callable, Iterable

import pythoncom
import win32com.client


LogCallback = Callable[[str], None]


def _log(message: str, callback: LogCallback | None = None) -> None:
    if callback:
        callback(message)
    else:
        print(message)


def _iter_files(folder: str, extensions: Iterable[str]) -> list[str]:
    return [
        file
        for file in os.listdir(folder)
        if file.lower().endswith(tuple(extensions)) and not file.startswith("~")
    ]


def batch_ppt_to_pdf(
    folder: str,
    delete_source: bool = True,
    log_callback: LogCallback | None = None,
) -> int:
    """Convert PowerPoint files in a folder to PDF."""
    pythoncom.CoInitialize()
    ppt_files = _iter_files(folder, (".ppt", ".pptx"))
    if not ppt_files:
        _log("未找到 PPT 文件。", log_callback)
        pythoncom.CoUninitialize()
        return 0

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1
    converted_files: list[str] = []
    max_retries = 3

    try:
        _log(f"找到 {len(ppt_files)} 个 PPT 文件，开始转换。", log_callback)

        for retry_count in range(max_retries):
            remaining_files = [file for file in ppt_files if file not in converted_files]
            if not remaining_files:
                break

            if retry_count > 0:
                _log(f"第 {retry_count} 次重试，继续转换剩余文件。", log_callback)

            for file in remaining_files:
                full_path = os.path.join(folder, file)
                pdf_path = os.path.join(folder, f"{os.path.splitext(file)[0]}.pdf")
                _log(f"正在转换 PPT：{file}", log_callback)

                try:
                    presentation = powerpoint.Presentations.Open(full_path)
                    presentation.SaveAs(pdf_path, 32)
                    presentation.Close()

                    if delete_source:
                        os.remove(full_path)
                        _log(f"已删除源文件：{file}", log_callback)

                    converted_files.append(file)
                    _log(f"完成：{os.path.basename(pdf_path)}", log_callback)
                except Exception as exc:
                    _log(f"转换失败：{file}；原因：{exc}", log_callback)
                    try:
                        powerpoint.Quit()
                    except Exception:
                        pass
                    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                    powerpoint.Visible = 1
    finally:
        try:
            powerpoint.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    _log(f"PPT 转换完成，共成功转换 {len(converted_files)} 个文件。", log_callback)
    return len(converted_files)


def batch_doc_to_pdf(
    folder: str,
    delete_source: bool = True,
    log_callback: LogCallback | None = None,
) -> int:
    """Convert Word files in a folder to PDF."""
    pythoncom.CoInitialize()
    doc_files = _iter_files(folder, (".doc", ".docx"))
    if not doc_files:
        _log("未找到 Word 文件。", log_callback)
        pythoncom.CoUninitialize()
        return 0

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = 1
    converted_files: list[str] = []
    max_retries = 3

    try:
        _log(f"找到 {len(doc_files)} 个 Word 文件，开始转换。", log_callback)

        for retry_count in range(max_retries):
            remaining_files = [file for file in doc_files if file not in converted_files]
            if not remaining_files:
                break

            if retry_count > 0:
                _log(f"第 {retry_count} 次重试，继续转换剩余文件。", log_callback)

            for file in remaining_files:
                full_path = os.path.join(folder, file)
                pdf_path = os.path.join(folder, f"{os.path.splitext(file)[0]}.pdf")
                _log(f"正在转换 Word：{file}", log_callback)

                try:
                    doc = word.Documents.Open(full_path)
                    doc.SaveAs(pdf_path, 17)
                    doc.Close()

                    if delete_source:
                        os.remove(full_path)
                        _log(f"已删除源文件：{file}", log_callback)

                    converted_files.append(file)
                    _log(f"完成：{os.path.basename(pdf_path)}", log_callback)
                except Exception as exc:
                    _log(f"转换失败：{file}；原因：{exc}", log_callback)
                    try:
                        word.Quit()
                    except Exception:
                        pass
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = 1
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    _log(f"Word 转换完成，共成功转换 {len(converted_files)} 个文件。", log_callback)
    return len(converted_files)


def flatten_folders(folder: str, log_callback: LogCallback | None = None) -> tuple[int, int]:
    """Move files from subfolders into the selected folder and remove empty folders."""
    _log(f"开始整理子文件夹：{folder}", log_callback)

    files_to_move: list[str] = []
    for root, _dirs, files in os.walk(folder, topdown=False):
        if root == folder:
            continue

        for file in files:
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
        f"子文件夹整理完成，共移动 {moved_count} 个文件，删除 {deleted_count} 个空文件夹。",
        log_callback,
    )
    return moved_count, deleted_count


def convert_folder(
    folder: str,
    flatten: bool = False,
    delete_source: bool = True,
    log_callback: LogCallback | None = None,
) -> tuple[int, int]:
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"文件夹不存在：{folder}")

    if flatten:
        flatten_folders(folder, log_callback)

    ppt_count = batch_ppt_to_pdf(folder, delete_source, log_callback)
    doc_count = batch_doc_to_pdf(folder, delete_source, log_callback)
    _log(f"全部完成：PPT {ppt_count} 个，Word {doc_count} 个。", log_callback)
    return ppt_count, doc_count


def main() -> int:
    folder_path = input("请输入需要批量转换的文件夹路径：").strip().strip('"')
    if not os.path.exists(folder_path):
        print(f"错误：文件夹不存在：{folder_path}")
        return 1

    flatten_choice = input("是否整理子文件夹，把里面的文件移动到主文件夹？(y/n)：").lower()
    delete_choice = input("转换成功后是否删除源文件？(y/n)：").lower()
    convert_folder(
        folder_path,
        flatten=flatten_choice == "y",
        delete_source=delete_choice == "y",
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
