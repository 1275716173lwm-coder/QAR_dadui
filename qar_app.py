from __future__ import annotations

import queue
import threading
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from qar_analyzer import AnalysisResult, LogRecord, ProgressEvent, analyze_selected_folder
from qar_output import OutputError, write_log, write_workbook


@dataclass
class WorkerMessage:
    kind: str
    payload: object


class QarApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.base_dir = Path(__file__).resolve().parent
        self.template_path = self.base_dir / "result_sample" / "result_sample.xlsx"
        self.output_dir = self.base_dir / "result"
        self.messages: queue.Queue[WorkerMessage] = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.close_requested = False

        root.title("QAR批量分析")
        root.geometry("760x520")
        root.minsize(680, 440)
        root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.folder_var = tk.StringVar()
        self.person_var = tk.StringVar(value="尚未开始")
        self.file_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="请选择包含人员子文件夹的数据目录")

        self._build_ui()
        self.root.after(100, self._poll_messages)

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="数据目录").grid(row=0, column=0, sticky="w")
        self.folder_entry = ttk.Entry(container, textvariable=self.folder_var)
        self.folder_entry.grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self.browse_button = ttk.Button(container, text="选择文件夹", command=self._choose_folder)
        self.browse_button.grid(row=1, column=1, sticky="ew")

        buttons = ttk.Frame(container)
        buttons.grid(row=2, column=0, columnspan=2, sticky="w", pady=(12, 14))
        self.start_button = ttk.Button(buttons, text="开始分析", command=self._start)
        self.start_button.pack(side="left")
        self.cancel_button = ttk.Button(buttons, text="取消", command=self._cancel, state="disabled")
        self.cancel_button.pack(side="left", padx=(8, 0))

        info = ttk.LabelFrame(container, text="分析进度", padding=10)
        info.grid(row=3, column=0, columnspan=2, sticky="nsew")
        info.columnconfigure(1, weight=1)
        info.rowconfigure(4, weight=1)

        ttk.Label(info, text="当前人员：").grid(row=0, column=0, sticky="nw")
        ttk.Label(info, textvariable=self.person_var).grid(row=0, column=1, sticky="nw")
        ttk.Label(info, text="当前文件：").grid(row=1, column=0, sticky="nw")
        ttk.Label(info, textvariable=self.file_var, wraplength=580).grid(row=1, column=1, sticky="nw")
        self.progress = ttk.Progressbar(info, mode="determinate", maximum=1)
        self.progress.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 6))
        ttk.Label(info, textvariable=self.status_var).grid(row=3, column=0, columnspan=2, sticky="w")

        self.log_text = tk.Text(info, height=12, state="disabled", wrap="word")
        self.log_text.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        scrollbar = ttk.Scrollbar(info, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=4, column=2, sticky="ns", pady=(8, 0))
        self.log_text.configure(yscrollcommand=scrollbar.set)

        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

    def _choose_folder(self) -> None:
        selected = filedialog.askdirectory(title="选择包含人员子文件夹的数据目录")
        if selected:
            self.folder_var.set(selected)

    def _append_status(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_running(self, running: bool) -> None:
        normal = "disabled" if running else "normal"
        self.start_button.configure(state=normal)
        self.browse_button.configure(state=normal)
        self.folder_entry.configure(state=normal)
        self.cancel_button.configure(state="normal" if running else "disabled")

    def _start(self) -> None:
        folder_text = self.folder_var.get().strip()
        if not folder_text:
            messagebox.showwarning("未选择目录", "请先选择包含人员子文件夹的数据目录。")
            return
        selected_folder = Path(folder_text)
        if not selected_folder.exists() or not selected_folder.is_dir():
            messagebox.showerror("目录无效", "所选路径不存在或不是文件夹。")
            return
        if not self.template_path.exists():
            messagebox.showerror("模板缺失", f"找不到结果模板：\n{self.template_path}")
            return

        self.cancel_event.clear()
        self.close_requested = False
        self.progress.configure(value=0, maximum=1)
        self.person_var.set("准备中")
        self.file_var.set("")
        self.status_var.set("正在扫描数据文件……")
        self._append_status(f"开始分析：{selected_folder}")
        self._set_running(True)
        self.worker = threading.Thread(
            target=self._run_worker,
            args=(selected_folder,),
            daemon=False,
            name="qar-analysis-worker",
        )
        self.worker.start()

    def _progress_callback(self, event: ProgressEvent) -> None:
        self.messages.put(WorkerMessage("progress", event))

    def _run_worker(self, selected_folder: Path) -> None:
        result: AnalysisResult | None = None
        try:
            result = analyze_selected_folder(selected_folder, self.cancel_event, self._progress_callback)
            workbook_path: Path | None = None
            output_error: str | None = None
            if not result.cancelled:
                try:
                    workbook_path = write_workbook(self.template_path, self.output_dir, result.summaries)
                except OutputError as exc:
                    output_error = str(exc)
                    result.log_records.append(
                        LogRecord(
                            record_type="OUTPUT", timestamp=datetime.now(), status="FAILED",
                            error_code="WORKBOOK_OUTPUT_ERROR", message=output_error,
                        )
                    )
            log_path = write_log(self.output_dir, result)
            if result.cancelled:
                self.messages.put(WorkerMessage("cancelled", (result, log_path)))
            elif output_error:
                self.messages.put(WorkerMessage("error", (output_error, log_path)))
            else:
                self.messages.put(WorkerMessage("done", (result, workbook_path, log_path)))
        except Exception as exc:
            now = datetime.now()
            message = f"程序发生未处理错误：{type(exc).__name__}: {exc}"
            if result is None:
                result = AnalysisResult(
                    selected_folder=selected_folder,
                    started_at=now,
                    ended_at=now,
                    summaries=[],
                    flights=[],
                    log_records=[
                        LogRecord(
                            record_type="RUN", timestamp=now, status="FAILED",
                            error_code="FATAL_ERROR", message=message,
                        )
                    ],
                )
            try:
                log_path = write_log(self.output_dir, result)
                detail = f"{message}\n日志：{log_path}"
            except Exception as log_exc:
                detail = f"{message}\n同时无法写入日志：{log_exc}"
            self.messages.put(WorkerMessage("fatal", detail))

    def _cancel(self) -> None:
        if self.worker and self.worker.is_alive():
            self.cancel_event.set()
            self.cancel_button.configure(state="disabled")
            self.status_var.set("正在安全取消，请稍候……")
            self._append_status("已请求取消，将在当前安全检查点停止。")

    def _handle_progress(self, event: ProgressEvent) -> None:
        if event.person:
            self.person_var.set(event.person)
        if event.file_name:
            self.file_var.set(event.file_name)
        maximum = max(event.total, 1)
        self.progress.configure(maximum=maximum, value=min(event.completed, maximum))
        if event.message:
            self.status_var.set(event.message)
            if event.stage in {"started", "person", "progress", "analyzed", "cancelled"}:
                self._append_status(event.message)

    def _finish_ui(self) -> None:
        self._set_running(False)
        self.worker = None
        if self.close_requested:
            self.root.destroy()

    def _poll_messages(self) -> None:
        try:
            while True:
                message = self.messages.get_nowait()
                if message.kind == "progress":
                    self._handle_progress(message.payload)  # type: ignore[arg-type]
                elif message.kind == "done":
                    result, workbook_path, log_path = message.payload  # type: ignore[misc]
                    self.progress.configure(value=max(self.progress["maximum"], 1))
                    text = f"分析完成。\n结果表：{workbook_path}\n日志：{log_path}"
                    self.status_var.set("分析完成")
                    self._append_status(text.replace("\n", "；"))
                    self._finish_ui()
                    if not self.close_requested:
                        messagebox.showinfo("分析完成", text)
                elif message.kind == "cancelled":
                    _, log_path = message.payload  # type: ignore[misc]
                    text = f"分析已取消，原结果表未被替换。\n日志：{log_path}"
                    self.status_var.set("分析已取消")
                    self._append_status(text.replace("\n", "；"))
                    self._finish_ui()
                    if not self.close_requested:
                        messagebox.showinfo("已取消", text)
                elif message.kind == "error":
                    error_text, log_path = message.payload  # type: ignore[misc]
                    text = f"分析已完成，但结果表保存失败：\n{error_text}\n日志：{log_path}"
                    self.status_var.set("结果表保存失败")
                    self._append_status(text.replace("\n", "；"))
                    self._finish_ui()
                    if not self.close_requested:
                        messagebox.showerror("保存失败", text)
                elif message.kind == "fatal":
                    text = str(message.payload)
                    self.status_var.set("分析失败")
                    self._append_status(text.replace("\n", "；"))
                    self._finish_ui()
                    if not self.close_requested:
                        messagebox.showerror("分析失败", text)
        except queue.Empty:
            pass
        try:
            if self.root.winfo_exists():
                self.root.after(100, self._poll_messages)
        except tk.TclError:
            return

    def _on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not self.close_requested:
                self.close_requested = True
                self.cancel_event.set()
                self.cancel_button.configure(state="disabled")
                self.status_var.set("正在安全取消，完成后窗口将关闭……")
                self._append_status("关闭请求已收到，正在安全取消。")
            return
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    QarApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
