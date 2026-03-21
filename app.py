from __future__ import annotations

import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


class Md2DocxApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Md2Docx - Markdown to Word")
        self.geometry("840x560")
        self.minsize(760, 520)

        self.mode_var = tk.StringVar(value="single")
        self.input_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=True)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.is_running = False

        self._build_ui()
        self._apply_startup_args()
        self.after(120, self._drain_log_queue)

    def _build_ui(self) -> None:
        root = tk.Frame(self, padx=14, pady=14)
        root.pack(fill=tk.BOTH, expand=True)

        mode_frame = tk.LabelFrame(root, text="Mode", padx=10, pady=8)
        mode_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Radiobutton(
            mode_frame,
            text="Single Markdown file",
            variable=self.mode_var,
            value="single",
            command=self._on_mode_change,
        ).pack(anchor="w")
        tk.Radiobutton(
            mode_frame,
            text="Folder batch convert",
            variable=self.mode_var,
            value="folder",
            command=self._on_mode_change,
        ).pack(anchor="w")

        path_frame = tk.LabelFrame(root, text="Paths", padx=10, pady=10)
        path_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(path_frame, text="Input:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.input_entry = tk.Entry(path_frame, textvariable=self.input_path_var)
        self.input_entry.grid(row=0, column=1, sticky="ew", pady=5)
        self.input_btn = tk.Button(path_frame, text="Browse", command=self._pick_input)
        self.input_btn.grid(row=0, column=2, padx=(8, 0), pady=5)

        tk.Label(path_frame, text="Output dir:").grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=5,
        )
        self.output_entry = tk.Entry(path_frame, textvariable=self.output_dir_var)
        self.output_entry.grid(row=1, column=1, sticky="ew", pady=5)
        self.output_btn = tk.Button(path_frame, text="Browse", command=self._pick_output_dir)
        self.output_btn.grid(row=1, column=2, padx=(8, 0), pady=5)

        tk.Label(path_frame, text="Template:").grid(
            row=2,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=5,
        )
        self.template_entry = tk.Entry(path_frame, textvariable=self.template_var)
        self.template_entry.grid(row=2, column=1, sticky="ew", pady=5)
        self.template_btn = tk.Button(path_frame, text="Browse", command=self._pick_template)
        self.template_btn.grid(row=2, column=2, padx=(8, 0), pady=5)

        self.clear_template_btn = tk.Button(
            path_frame,
            text="Clear",
            command=lambda: self.template_var.set(""),
            width=8,
        )
        self.clear_template_btn.grid(row=2, column=3, padx=(8, 0), pady=5)

        path_frame.columnconfigure(1, weight=1)

        opt_frame = tk.LabelFrame(root, text="Options", padx=10, pady=10)
        opt_frame.pack(fill=tk.X, pady=(0, 10))

        self.recursive_check = tk.Checkbutton(
            opt_frame,
            text="Include subfolders when converting a folder",
            variable=self.recursive_var,
        )
        self.recursive_check.pack(anchor="w")

        action_frame = tk.Frame(root)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        self.convert_btn = tk.Button(
            action_frame,
            text="Convert",
            width=16,
            command=self._start_convert,
            bg="#1f6feb",
            fg="white",
        )
        self.convert_btn.pack(side=tk.LEFT)

        self.log_box = ScrolledText(root, height=14, wrap=tk.WORD)
        self.log_box.pack(fill=tk.BOTH, expand=True)
        self.log_box.insert(
            tk.END,
            "Ready. Choose input and output path, then click Convert.\n",
        )
        self.log_box.configure(state=tk.DISABLED)

        self._on_mode_change()

    def _on_mode_change(self) -> None:
        folder_mode = self.mode_var.get() == "folder"
        self.recursive_check.configure(state=(tk.NORMAL if folder_mode else tk.DISABLED))

    def _apply_startup_args(self) -> None:
        if len(sys.argv) < 2:
            return

        arg_path = Path(sys.argv[1]).expanduser()
        if not arg_path.exists():
            return

        if arg_path.is_file():
            self.mode_var.set("single")
            self.input_path_var.set(str(arg_path))
            if not self.output_dir_var.get():
                self.output_dir_var.set(str(arg_path.parent / "docx_output"))
        elif arg_path.is_dir():
            self.mode_var.set("folder")
            self.input_path_var.set(str(arg_path))
            if not self.output_dir_var.get():
                self.output_dir_var.set(str(arg_path / "docx_output"))

        self._on_mode_change()

    def _pick_input(self) -> None:
        if self.mode_var.get() == "single":
            path = filedialog.askopenfilename(
                title="Choose Markdown file",
                filetypes=[("Markdown", "*.md *.markdown *.mdown *.mkd"), ("All files", "*.*")],
            )
        else:
            path = filedialog.askdirectory(title="Choose source folder")

        if path:
            self.input_path_var.set(path)

            if not self.output_dir_var.get():
                path_obj = Path(path)
                if self.mode_var.get() == "single":
                    self.output_dir_var.set(str(path_obj.parent / "docx_output"))
                else:
                    self.output_dir_var.set(str(path_obj / "docx_output"))

    def _pick_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Choose output folder")
        if path:
            self.output_dir_var.set(path)

    def _pick_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose reference docx (optional)",
            filetypes=[("Word docx", "*.docx")],
        )
        if path:
            self.template_var.set(path)

    def _resolve_pandoc_path(self) -> str | None:
        pandoc_name = "pandoc.exe" if sys.platform.startswith("win") else "pandoc"
        candidates: list[Path] = []

        # PyInstaller onefile: bundled files are extracted to _MEIPASS.
        if getattr(sys, "frozen", False):
            meipass = getattr(sys, "_MEIPASS", None)
            if meipass:
                candidates.append(Path(meipass) / pandoc_name)

            exe_dir = Path(sys.executable).resolve().parent
            candidates.append(exe_dir / pandoc_name)

        script_dir = Path(__file__).resolve().parent
        candidates.append(script_dir / pandoc_name)
        candidates.append(script_dir / "vendor" / pandoc_name)

        for candidate in candidates:
            if candidate.is_file():
                return str(candidate)

        return shutil.which("pandoc")

    def _start_convert(self) -> None:
        if self.is_running:
            return

        pandoc_path = self._resolve_pandoc_path()
        if not pandoc_path:
            messagebox.showerror(
                "Pandoc Not Found",
                "No bundled pandoc found, and pandoc is not available in PATH.",
            )
            return

        input_path = self.input_path_var.get().strip()
        output_dir = self.output_dir_var.get().strip()
        template = self.template_var.get().strip()

        if not input_path:
            messagebox.showwarning("Missing Input", "Please choose an input file/folder.")
            return
        if not output_dir:
            messagebox.showwarning("Missing Output", "Please choose an output folder.")
            return

        in_path = Path(input_path)
        out_dir = Path(output_dir)

        if self.mode_var.get() == "single":
            if not in_path.is_file():
                messagebox.showerror("Invalid Input", "Input should be a Markdown file.")
                return
            if in_path.suffix.lower() not in {".md", ".markdown", ".mdown", ".mkd"}:
                messagebox.showerror("Invalid Input", "Selected file is not a Markdown file.")
                return
        else:
            if not in_path.is_dir():
                messagebox.showerror("Invalid Input", "Input should be a folder.")
                return

        if template:
            template_path = Path(template)
            if not template_path.is_file() or template_path.suffix.lower() != ".docx":
                messagebox.showerror("Invalid Template", "Template must be an existing .docx file.")
                return

        self._set_running(True)
        self._append_log("-" * 72)
        self._append_log(f"Using pandoc: {pandoc_path}")
        self._append_log(f"Mode: {self.mode_var.get()}")
        self._append_log(f"Input: {in_path}")
        self._append_log(f"Output: {out_dir}")
        self._append_log(f"Template: {template if template else '(none)'}")

        mode = self.mode_var.get()
        recursive = self.recursive_var.get()

        thread = threading.Thread(
            target=self._convert_worker,
            args=(in_path, out_dir, template, mode, recursive, pandoc_path),
            daemon=True,
        )
        thread.start()

    def _convert_worker(
        self,
        in_path: Path,
        out_dir: Path,
        template: str,
        mode: str,
        recursive: bool,
        pandoc_path: str,
    ) -> None:
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
            md_files = self._collect_md_files(in_path, mode, recursive)

            if not md_files:
                self.log_queue.put("No markdown files found.")
                return

            self.log_queue.put(f"Found {len(md_files)} markdown file(s).")

            ok = 0
            fail = 0

            for idx, src in enumerate(md_files, start=1):
                if mode == "single":
                    dst = out_dir / f"{src.stem}.docx"
                else:
                    rel = src.relative_to(in_path)
                    dst = out_dir / rel.with_suffix(".docx")

                dst.parent.mkdir(parents=True, exist_ok=True)
                self.log_queue.put(f"[{idx}/{len(md_files)}] {src} -> {dst}")

                try:
                    self._run_pandoc(src, dst, template, pandoc_path)
                    ok += 1
                    self.log_queue.put("  OK")
                except Exception as exc:
                    fail += 1
                    self.log_queue.put(f"  FAIL: {exc}")

            self.log_queue.put("=" * 72)
            self.log_queue.put(f"Done. Success: {ok}, Failed: {fail}")
        finally:
            self.log_queue.put("__UI_UNLOCK__")

    def _collect_md_files(self, in_path: Path, mode: str, recursive: bool) -> list[Path]:
        if mode == "single":
            return [in_path]

        suffixes = {".md", ".markdown", ".mdown", ".mkd"}
        pattern = "**/*" if recursive else "*"
        files = sorted(in_path.glob(pattern))
        return [f for f in files if f.is_file() and f.suffix.lower() in suffixes]

    def _run_pandoc(self, src: Path, dst: Path, template: str, pandoc_path: str) -> None:
        cmd = [
            pandoc_path,
            str(src),
            "-o",
            str(dst),
            "--resource-path",
            str(src.parent),
        ]

        if template:
            cmd.extend(["--reference-doc", template])

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            err = result.stderr.strip() or result.stdout.strip() or "pandoc conversion failed"
            raise RuntimeError(err)

    def _set_running(self, running: bool) -> None:
        self.is_running = running
        state = tk.DISABLED if running else tk.NORMAL

        self.convert_btn.configure(state=state, text=("Converting..." if running else "Convert"))
        self.input_btn.configure(state=state)
        self.output_btn.configure(state=state)
        self.template_btn.configure(state=state)
        self.clear_template_btn.configure(state=state)
        self.input_entry.configure(state=state)
        self.output_entry.configure(state=state)
        self.template_entry.configure(state=state)

    def _append_log(self, text: str) -> None:
        self.log_box.configure(state=tk.NORMAL)
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)
        self.log_box.configure(state=tk.DISABLED)

    def _drain_log_queue(self) -> None:
        try:
            while True:
                msg = self.log_queue.get_nowait()
                if msg == "__UI_UNLOCK__":
                    self._set_running(False)
                else:
                    self._append_log(msg)
        except queue.Empty:
            pass
        finally:
            self.after(120, self._drain_log_queue)


if __name__ == "__main__":
    app = Md2DocxApp()
    app.mainloop()
