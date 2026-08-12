import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from run import convert_folder


class PdfConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PPT / Word 批量转 PDF")
        self.geometry("760x560")
        self.minsize(680, 480)

        self.folder_var = tk.StringVar()
        self.flatten_var = tk.BooleanVar(value=False)
        self.delete_source_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="请选择要处理的文件夹")
        self.log_queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker: threading.Thread | None = None

        self._build_ui()
        self.after(100, self._drain_log_queue)

    def _build_ui(self) -> None:
        self.configure(bg="#f6f7f9")

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f6f7f9")
        style.configure("TLabel", background="#f6f7f9", font=("Microsoft YaHei UI", 10))
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Hint.TLabel", foreground="#5f6b7a")
        style.configure("TButton", font=("Microsoft YaHei UI", 10), padding=(14, 8))
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("TCheckbutton", background="#f6f7f9", font=("Microsoft YaHei UI", 10))

        container = ttk.Frame(self, padding=24)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="批量转 PDF", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            container,
            text="选择一个文件夹，一键把里面的 PPT、PPTX、DOC、DOCX 转成 PDF。",
            style="Hint.TLabel",
        ).pack(anchor=tk.W, pady=(6, 20))

        path_frame = ttk.Frame(container)
        path_frame.pack(fill=tk.X)

        self.path_entry = ttk.Entry(path_frame, textvariable=self.folder_var, font=("Microsoft YaHei UI", 10))
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)
        ttk.Button(path_frame, text="选择文件夹", command=self._choose_folder).pack(side=tk.LEFT, padx=(10, 0))

        options_frame = ttk.Frame(container)
        options_frame.pack(fill=tk.X, pady=(18, 8))

        ttk.Checkbutton(
            options_frame,
            text="先整理子文件夹：把子文件夹里的文件移动到当前文件夹",
            variable=self.flatten_var,
        ).pack(anchor=tk.W, pady=3)
        ttk.Checkbutton(
            options_frame,
            text="转换成功后删除原 PPT / Word 文件",
            variable=self.delete_source_var,
        ).pack(anchor=tk.W, pady=3)

        actions_frame = ttk.Frame(container)
        actions_frame.pack(fill=tk.X, pady=(12, 12))

        self.start_button = ttk.Button(
            actions_frame,
            text="开始转换",
            style="Accent.TButton",
            command=self._start_convert,
        )
        self.start_button.pack(side=tk.LEFT)
        ttk.Button(actions_frame, text="打开输出文件夹", command=self._open_folder).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Label(actions_frame, textvariable=self.status_var, style="Hint.TLabel").pack(side=tk.LEFT, padx=(16, 0))

        self.progress = ttk.Progressbar(container, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=(0, 14))

        ttk.Label(container, text="处理记录").pack(anchor=tk.W)
        self.log_text = scrolledtext.ScrolledText(
            container,
            height=14,
            wrap=tk.WORD,
            font=("Consolas", 10),
            relief=tk.FLAT,
            borderwidth=1,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.log_text.configure(state=tk.DISABLED)

    def _choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择要转换的文件夹")
        if folder:
            self.folder_var.set(folder)
            self.status_var.set("已选择文件夹")

    def _open_folder(self) -> None:
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择一个有效的文件夹。")
            return
        os.startfile(folder)

    def _start_convert(self) -> None:
        folder = self.folder_var.get().strip().strip('"')
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择一个有效的文件夹。")
            return

        if self.worker and self.worker.is_alive():
            return

        self._set_running(True)
        self._clear_log()
        self._append_log("开始处理，请不要关闭 PowerPoint 或 Word 窗口。")

        self.worker = threading.Thread(
            target=self._convert_in_background,
            args=(folder, self.flatten_var.get(), self.delete_source_var.get()),
            daemon=True,
        )
        self.worker.start()

    def _convert_in_background(self, folder: str, flatten: bool, delete_source: bool) -> None:
        try:
            ppt_count, doc_count = convert_folder(
                folder,
                flatten=flatten,
                delete_source=delete_source,
                log_callback=lambda message: self.log_queue.put(("log", message)),
            )
            self.log_queue.put(("done", f"转换完成：PPT {ppt_count} 个，Word {doc_count} 个。"))
        except Exception as exc:
            self.log_queue.put(("error", str(exc)))

    def _drain_log_queue(self) -> None:
        try:
            while True:
                kind, message = self.log_queue.get_nowait()
                if kind == "log":
                    self._append_log(message)
                elif kind == "done":
                    self._append_log(message)
                    self._set_running(False)
                    messagebox.showinfo("完成", message)
                elif kind == "error":
                    self._append_log(f"处理失败：{message}")
                    self._set_running(False)
                    messagebox.showerror("处理失败", message)
        except queue.Empty:
            pass

        self.after(100, self._drain_log_queue)

    def _set_running(self, running: bool) -> None:
        if running:
            self.start_button.configure(state=tk.DISABLED)
            self.progress.start(10)
            self.status_var.set("正在转换...")
        else:
            self.start_button.configure(state=tk.NORMAL)
            self.progress.stop()
            self.status_var.set("处理完成")

    def _clear_log(self) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)


if __name__ == "__main__":
    app = PdfConverterApp()
    app.mainloop()
