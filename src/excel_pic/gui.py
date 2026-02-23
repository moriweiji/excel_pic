from __future__ import annotations

import os
import subprocess
import sys
import threading
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from . import (
    AppError,
    DEFAULT_PREFIX_TEXT,
    detect_runtime_dir,
    get_output_paths,
    run_generation,
)


class ExcelPicApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("剧集 Excel 生成器")
        self.root.geometry("980x760")

        self.exe_dir = detect_runtime_dir()

        self.images_dir_var = tk.StringVar()
        self.word_file_var = tk.StringVar()
        self.output_dir_var = tk.StringVar(value="-")
        self.logs_dir_var = tk.StringVar(value=str(self.exe_dir / "data" / "logs"))
        self.strict_var = tk.BooleanVar(value=False)

        self.last_export_dir: Path | None = None
        self.last_report_path: Path | None = None

        self._build_ui()
        self._bind_events()
        self._load_template_on_start()

    def _build_ui(self) -> None:
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(7, weight=1)

        title = tk.Label(
            self.root,
            text="剧集 Excel 生成器（V1）",
            font=("Microsoft YaHei", 16, "bold"),
            anchor="w",
        )
        title.grid(row=0, column=0, columnspan=3, sticky="ew", padx=12, pady=(12, 4))

        desc = tk.Label(
            self.root,
            text="选择图片目录和Word文件，自动生成导出目录、Excel和日志",
            anchor="w",
            fg="#555",
        )
        desc.grid(row=1, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 12))

        tk.Label(self.root, text="图片源目录").grid(row=2, column=0, sticky="w", padx=12, pady=4)
        tk.Entry(self.root, textvariable=self.images_dir_var).grid(row=2, column=1, sticky="ew", padx=6, pady=4)
        tk.Button(self.root, text="浏览...", command=self._choose_images_dir).grid(row=2, column=2, sticky="ew", padx=12, pady=4)

        tk.Label(self.root, text="Word 文件").grid(row=3, column=0, sticky="w", padx=12, pady=4)
        tk.Entry(self.root, textvariable=self.word_file_var).grid(row=3, column=1, sticky="ew", padx=6, pady=4)
        tk.Button(self.root, text="浏览...", command=self._choose_word_file).grid(row=3, column=2, sticky="ew", padx=12, pady=4)

        tk.Label(self.root, text="通用前缀模板（每个分景自动插入）").grid(
            row=4, column=0, columnspan=3, sticky="w", padx=12, pady=(10, 4)
        )
        self.prefix_text = ScrolledText(self.root, height=10, wrap=tk.WORD)
        self.prefix_text.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=12, pady=4)

        btn_row = tk.Frame(self.root)
        btn_row.grid(row=6, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 8))
        self.btn_load_tpl = tk.Button(btn_row, text="加载模板", command=self._load_template)
        self.btn_load_tpl.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_save_tpl = tk.Button(btn_row, text="保存模板", command=self._save_template)
        self.btn_save_tpl.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_reset_tpl = tk.Button(btn_row, text="恢复默认", command=self._reset_template)
        self.btn_reset_tpl.pack(side=tk.LEFT)

        info_frame = tk.Frame(self.root)
        info_frame.grid(row=7, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 6))
        info_frame.grid_columnconfigure(1, weight=1)

        tk.Label(info_frame, text="输出目录").grid(row=0, column=0, sticky="w", pady=2)
        tk.Entry(info_frame, textvariable=self.output_dir_var, state="readonly").grid(
            row=0, column=1, sticky="ew", padx=6, pady=2
        )

        tk.Label(info_frame, text="日志目录").grid(row=1, column=0, sticky="w", pady=2)
        tk.Entry(info_frame, textvariable=self.logs_dir_var, state="readonly").grid(
            row=1, column=1, sticky="ew", padx=6, pady=2
        )

        self.chk_strict = tk.Checkbutton(info_frame, text="严格模式（坏图等问题直接中断）", variable=self.strict_var)
        self.chk_strict.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        action_frame = tk.Frame(self.root)
        action_frame.grid(row=8, column=0, columnspan=3, sticky="ew", padx=12, pady=(4, 8))
        self.btn_start = tk.Button(action_frame, text="开始生成", command=self._start_generate)
        self.btn_start.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_clear_log = tk.Button(action_frame, text="清空日志", command=self._clear_log)
        self.btn_clear_log.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_open_export = tk.Button(action_frame, text="打开输出目录", command=self._open_last_export)
        self.btn_open_export.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_open_logs = tk.Button(action_frame, text="打开日志目录", command=self._open_logs_dir)
        self.btn_open_logs.pack(side=tk.LEFT)

        tk.Label(self.root, text="运行日志").grid(row=9, column=0, columnspan=3, sticky="w", padx=12)
        self.log_text = ScrolledText(self.root, height=12, wrap=tk.WORD)
        self.log_text.grid(row=10, column=0, columnspan=3, sticky="nsew", padx=12, pady=(4, 12))
        self.root.grid_rowconfigure(10, weight=1)

    def _bind_events(self) -> None:
        self.images_dir_var.trace_add("write", lambda *_: self._refresh_output_preview())

    def _refresh_output_preview(self) -> None:
        images_dir = self.images_dir_var.get().strip()
        if not images_dir:
            self.output_dir_var.set("-")
            return
        try:
            _data, _cfg, _root, _logs, export_dir = get_output_paths(self.exe_dir, Path(images_dir))
            self.output_dir_var.set(str(export_dir))
        except Exception:
            self.output_dir_var.set("-")

    def _choose_images_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择图片源目录")
        if selected:
            self.images_dir_var.set(selected)

    def _choose_word_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="选择 Word 文件",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if selected:
            self.word_file_var.set(selected)

    def _template_file(self) -> Path:
        return self.exe_dir / "data" / "config" / "prompt_prefix.txt"

    def _load_template_on_start(self) -> None:
        self._load_template(show_tip=False)

    def _load_template(self, show_tip: bool = True) -> None:
        cfg_file = self._template_file()
        cfg_file.parent.mkdir(parents=True, exist_ok=True)
        if not cfg_file.exists():
            cfg_file.write_text(DEFAULT_PREFIX_TEXT, encoding="utf-8")
        text = cfg_file.read_text(encoding="utf-8")
        self.prefix_text.delete("1.0", tk.END)
        self.prefix_text.insert("1.0", text)
        if show_tip:
            self._append_log("INFO", f"已加载模板：{cfg_file}")

    def _save_template(self) -> None:
        text = self.prefix_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showerror("模板为空", "通用前缀模板不能为空，请填写后再保存。")
            return
        cfg_file = self._template_file()
        cfg_file.parent.mkdir(parents=True, exist_ok=True)
        cfg_file.write_text(text, encoding="utf-8")
        self._append_log("INFO", f"模板已保存：{cfg_file}")

    def _reset_template(self) -> None:
        self.prefix_text.delete("1.0", tk.END)
        self.prefix_text.insert("1.0", DEFAULT_PREFIX_TEXT)
        self._append_log("INFO", "模板已恢复默认，请按需保存。")

    def _append_log(self, level: str, message: str) -> None:
        self.log_text.insert(tk.END, f"[{level}] {message}\n")
        self.log_text.see(tk.END)

    def _clear_log(self) -> None:
        self.log_text.delete("1.0", tk.END)

    def _set_running(self, running: bool) -> None:
        state = tk.DISABLED if running else tk.NORMAL
        self.btn_start.config(state=state)
        self.btn_load_tpl.config(state=state)
        self.btn_save_tpl.config(state=state)
        self.btn_reset_tpl.config(state=state)

    def _confirm_overwrite_from_worker(self, export_dir: Path) -> bool:
        event = threading.Event()
        holder: dict[str, bool] = {"ok": False}

        def ask() -> None:
            holder["ok"] = messagebox.askyesno("覆盖确认", f"导出目录已存在：\n{export_dir}\n\n是否覆盖？")
            event.set()

        self.root.after(0, ask)
        event.wait()
        return holder["ok"]

    def _open_path(self, path: Path) -> None:
        if not path.exists():
            messagebox.showwarning("路径不存在", f"路径不存在：\n{path}")
            return

        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
            return
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
            return
        subprocess.run(["xdg-open", str(path)], check=False)

    def _open_last_export(self) -> None:
        if not self.last_export_dir:
            messagebox.showinfo("暂无输出", "请先执行生成。")
            return
        self._open_path(self.last_export_dir)

    def _open_logs_dir(self) -> None:
        logs_dir = self.exe_dir / "data" / "logs"
        self._open_path(logs_dir)

    def _start_generate(self) -> None:
        images_dir = self.images_dir_var.get().strip()
        word_file = self.word_file_var.get().strip()
        prefix_text = self.prefix_text.get("1.0", tk.END).strip()

        if not images_dir:
            messagebox.showerror("缺少输入", "请先选择图片源目录。")
            return
        if not word_file:
            messagebox.showerror("缺少输入", "请先选择 Word 文件。")
            return
        if not prefix_text:
            messagebox.showerror("模板为空", "通用前缀模板不能为空，请填写后再生成。")
            return

        self._set_running(True)
        self._append_log("INFO", "开始生成任务...")

        def worker() -> None:
            try:
                export_dir, excel_path, report_path = run_generation(
                    images_dir=Path(images_dir),
                    word_file=Path(word_file),
                    strict=self.strict_var.get(),
                    prefix_text_override=prefix_text,
                    exe_dir=self.exe_dir,
                    confirm_overwrite=self._confirm_overwrite_from_worker,
                    log_sink=lambda level, msg: self.root.after(0, self._append_log, level, msg),
                )
                self.last_export_dir = export_dir
                self.last_report_path = report_path
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "生成完成",
                        f"导出目录：\n{export_dir}\n\nExcel：\n{excel_path}\n\n报告：\n{report_path}",
                    ),
                )
            except AppError as exc:
                extra = ""
                if exc.log_path:
                    extra += f"\n\n日志：\n{exc.log_path}"
                if exc.report_path:
                    extra += f"\n\n报告：\n{exc.report_path}"
                self.root.after(0, lambda: messagebox.showerror("生成失败", f"[{exc.code}] {exc.message}{extra}"))
            finally:
                self.root.after(0, lambda: self._set_running(False))

        threading.Thread(target=worker, daemon=True).start()


def run() -> None:
    root = tk.Tk()
    ExcelPicApp(root)
    root.mainloop()


if __name__ == "__main__":
    run()
