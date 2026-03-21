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


TRANSLATIONS = {
    "zh": {
        "window_title": "Md2Docx - Markdown 转 Word",
        "header_title": "Markdown → Word",
        "header_subtitle": "现代化转换器 · 支持中英文切换、标准/高级双模式",
        "language_group": "Language / 语言",
        "edition_group": "界面模式",
        "settings_show_btn": "展开设置",
        "settings_hide_btn": "收起设置",
        "logs_show_btn": "展开日志",
        "logs_hide_btn": "收起日志",
        "logs_title": "日志",
        "edition_standard": "标准模式",
        "edition_advanced": "高级模式",
        "standard_title": "标准模式",
        "standard_mode_group": "标准子模式",
        "standard_mode_paste": "复制模式",
        "standard_mode_single": "单文件模式",
        "advanced_title": "高级模式",
        "advanced_mode_group": "高级子模式",
        "advanced_mode_single": "单文件模式",
        "advanced_mode_folder": "文件夹批量模式",
        "advanced_recursive": "批量时包含子目录",
        "input_label": "输入：",
        "output_label": "输出目录：",
        "filename_label": "文件名：",
        "template_label": "模板：",
        "markdown_content_label": "Markdown 内容：",
        "paste_hint": "提示：直接粘贴 Markdown 内容，支持标题、列表、表格等常见语法。",
        "paste_placeholder": "请将 Markdown 内容粘贴到这里，然后点击“转换”...",
        "browse_btn": "浏览",
        "clear_btn": "清空",
        "convert_btn": "转换",
        "converting_btn": "转换中...",
        "status_ready": "就绪",
        "status_running": "转换中，请稍候...",
        "status_success": "转换成功",
        "status_fail": "转换失败",
        "status_partial": "部分成功",
        "status_empty": "未执行转换",
        "log_ready": "已就绪：请选择输入并点击转换。",
        "mode_standard_paste": "标准-复制模式",
        "mode_standard_single": "标准-单文件模式",
        "mode_advanced_single": "高级-单文件模式",
        "mode_advanced_folder": "高级-文件夹模式",
        "choose_md_file": "选择 Markdown 文件",
        "choose_source_folder": "选择源目录",
        "choose_output_folder": "选择输出目录",
        "choose_template": "选择可选 Word 模板（.docx）",
        "pandoc_not_found_title": "未找到 Pandoc",
        "pandoc_not_found_msg": "未找到内置 pandoc，且 PATH 中也没有 pandoc。",
        "missing_input_title": "缺少输入",
        "missing_input_msg": "请选择输入文件/目录。",
        "missing_output_title": "缺少输出",
        "missing_output_msg": "请选择输出目录。",
        "missing_content_title": "缺少内容",
        "missing_content_msg": "请先粘贴 Markdown 内容。",
        "missing_filename_title": "缺少文件名",
        "missing_filename_msg": "请填写输出文件名。",
        "invalid_filename_title": "文件名无效",
        "invalid_filename_msg": "文件名不能包含以下字符：<>:\"/\\|?*",
        "invalid_input_title": "输入无效",
        "invalid_input_file_msg": "输入必须是 Markdown 文件。",
        "invalid_input_folder_msg": "输入必须是目录。",
        "invalid_template_title": "模板无效",
        "invalid_template_msg": "模板必须是存在的 .docx 文件。",
        "log_using_pandoc": "使用 pandoc：{pandoc}",
        "log_mode": "模式：{mode}",
        "log_input": "输入：{input}",
        "log_output": "输出：{output}",
        "log_template": "模板：{template}",
        "log_target": "目标：{target}",
        "none_text": "（无）",
        "log_found_files": "找到 {count} 个 Markdown 文件。",
        "log_no_md": "未找到 Markdown 文件。",
        "log_done": "完成。成功：{ok}，失败：{fail}",
        "log_ok": "  成功",
        "log_fail": "  失败：{err}",
        "result_success_title": "转换成功",
        "result_success_msg": "已成功生成 {ok} 个文件。",
        "result_fail_title": "转换失败",
        "result_fail_msg": "转换失败 {fail} 个文件。",
        "result_partial_title": "部分成功",
        "result_partial_msg": "成功 {ok} 个，失败 {fail} 个。",
        "result_empty_title": "未执行转换",
        "result_empty_msg": "没有可转换的内容或文件。",
    },
    "en": {
        "window_title": "Md2Docx - Markdown to Word",
        "header_title": "Markdown -> Word",
        "header_subtitle": "Modern converter with bilingual UI and Standard/Advanced workflows",
        "language_group": "Language / 语言",
        "edition_group": "Interface Mode",
        "settings_show_btn": "Show Settings",
        "settings_hide_btn": "Hide Settings",
        "logs_show_btn": "Show Logs",
        "logs_hide_btn": "Hide Logs",
        "logs_title": "Logs",
        "edition_standard": "Standard",
        "edition_advanced": "Advanced",
        "standard_title": "Standard Mode",
        "standard_mode_group": "Standard Submode",
        "standard_mode_paste": "Paste Mode",
        "standard_mode_single": "Single File Mode",
        "advanced_title": "Advanced Mode",
        "advanced_mode_group": "Advanced Submode",
        "advanced_mode_single": "Single File Mode",
        "advanced_mode_folder": "Folder Batch Mode",
        "advanced_recursive": "Include subfolders in folder mode",
        "input_label": "Input:",
        "output_label": "Output dir:",
        "filename_label": "File name:",
        "template_label": "Template:",
        "markdown_content_label": "Markdown content:",
        "paste_hint": "Tip: paste markdown directly. Headings, lists and tables are supported.",
        "paste_placeholder": "Paste Markdown content here, then click Convert...",
        "browse_btn": "Browse",
        "clear_btn": "Clear",
        "convert_btn": "Convert",
        "converting_btn": "Converting...",
        "status_ready": "Ready",
        "status_running": "Converting, please wait...",
        "status_success": "Success",
        "status_fail": "Failed",
        "status_partial": "Partial Success",
        "status_empty": "No Conversion",
        "log_ready": "Ready. Choose input and click Convert.",
        "mode_standard_paste": "Standard-Paste",
        "mode_standard_single": "Standard-Single",
        "mode_advanced_single": "Advanced-Single",
        "mode_advanced_folder": "Advanced-Folder",
        "choose_md_file": "Choose Markdown file",
        "choose_source_folder": "Choose source folder",
        "choose_output_folder": "Choose output folder",
        "choose_template": "Choose optional Word template (.docx)",
        "pandoc_not_found_title": "Pandoc Not Found",
        "pandoc_not_found_msg": "No bundled pandoc found, and pandoc is not available in PATH.",
        "missing_input_title": "Missing Input",
        "missing_input_msg": "Please choose an input file/folder.",
        "missing_output_title": "Missing Output",
        "missing_output_msg": "Please choose an output folder.",
        "missing_content_title": "Missing Content",
        "missing_content_msg": "Please paste Markdown content first.",
        "missing_filename_title": "Missing File Name",
        "missing_filename_msg": "Please provide output file name.",
        "invalid_filename_title": "Invalid File Name",
        "invalid_filename_msg": "File name cannot contain: <>:\"/\\|?*",
        "invalid_input_title": "Invalid Input",
        "invalid_input_file_msg": "Input should be a Markdown file.",
        "invalid_input_folder_msg": "Input should be a folder.",
        "invalid_template_title": "Invalid Template",
        "invalid_template_msg": "Template must be an existing .docx file.",
        "log_using_pandoc": "Using pandoc: {pandoc}",
        "log_mode": "Mode: {mode}",
        "log_input": "Input: {input}",
        "log_output": "Output: {output}",
        "log_template": "Template: {template}",
        "log_target": "Target: {target}",
        "none_text": "(none)",
        "log_found_files": "Found {count} markdown file(s).",
        "log_no_md": "No markdown files found.",
        "log_done": "Done. Success: {ok}, Failed: {fail}",
        "log_ok": "  OK",
        "log_fail": "  FAIL: {err}",
        "result_success_title": "Conversion Succeeded",
        "result_success_msg": "Generated {ok} file(s) successfully.",
        "result_fail_title": "Conversion Failed",
        "result_fail_msg": "Failed to generate {fail} file(s).",
        "result_partial_title": "Partial Success",
        "result_partial_msg": "Success: {ok}, Failed: {fail}.",
        "result_empty_title": "No Conversion",
        "result_empty_msg": "No content or files were converted.",
    },
}

MD_SUFFIXES = {".md", ".markdown", ".mdown", ".mkd"}
INVALID_FILENAME_CHARS = set('<>:"/\\|?*')


class Md2DocxApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.colors = {
            "bg": "#eef2ff",
            "card": "#ffffff",
            "text": "#0f172a",
            "muted": "#64748b",
            "primary": "#2563eb",
            "primary_hover": "#1d4ed8",
            "border": "#cbd5e1",
            "button": "#e2e8f0",
            "button_hover": "#cbd5e1",
            "log_bg": "#0b1220",
            "log_fg": "#dbe7ff",
        }
        self.fonts = {
            "body": ("Segoe UI", 10),
            "heading": ("Segoe UI Semibold", 10),
            "title": ("Segoe UI Semibold", 18),
            "subtitle": ("Segoe UI", 10),
            "button": ("Segoe UI Semibold", 10),
            "mono": ("Consolas", 10),
        }

        self.configure(bg=self.colors["bg"])
        self.option_add("*Font", self.fonts["body"])
        self.geometry("1040x780")
        self.minsize(940, 680)

        self.lang_var = tk.StringVar(value="zh")
        self.edition_var = tk.StringVar(value="standard")
        self.standard_mode_var = tk.StringVar(value="paste")
        self.advanced_mode_var = tk.StringVar(value="single")

        self.std_paste_output_var = tk.StringVar()
        self.std_paste_filename_var = tk.StringVar(value="output")
        self.std_paste_template_var = tk.StringVar()

        self.std_single_input_var = tk.StringVar()
        self.std_single_output_var = tk.StringVar()
        self.std_single_template_var = tk.StringVar()

        self.adv_input_var = tk.StringVar()
        self.adv_output_var = tk.StringVar()
        self.adv_template_var = tk.StringVar()
        self.adv_recursive_var = tk.BooleanVar(value=True)

        self.log_queue: queue.Queue[object] = queue.Queue()
        self.is_running = False
        self.paste_placeholder_active = False
        self.settings_visible = False
        self.log_visible = False
        self.status_kind = "ready"

        self.interactive_widgets: list[tk.Widget] = []

        self._build_ui()
        self._apply_modern_theme()
        self._refresh_texts()
        self._setup_paste_placeholder()
        self._apply_startup_args()
        self._append_log(self.t("log_ready"))
        self._set_status("ready")
        self.after(120, self._drain_log_queue)

    def t(self, key: str, lang: str | None = None, **kwargs: object) -> str:
        use_lang = lang or self.lang_var.get()
        text = TRANSLATIONS.get(use_lang, TRANSLATIONS["en"])[key]
        return text.format(**kwargs) if kwargs else text

    def _build_ui(self) -> None:
        root = tk.Frame(self, padx=16, pady=16)
        root.pack(fill=tk.BOTH, expand=True)
        self.root_frame = root

        self.header_frame = tk.Frame(root)
        self.header_frame.pack(fill=tk.X, pady=(0, 12))
        self.header_title_label = tk.Label(self.header_frame, anchor="w")
        self.header_title_label.pack(fill=tk.X)
        self.header_subtitle_label = tk.Label(self.header_frame, anchor="w")
        self.header_subtitle_label.pack(fill=tk.X, pady=(2, 0))

        self.toolbar_frame = tk.Frame(root)
        self.toolbar_frame.pack(fill=tk.X, pady=(0, 10))

        self.settings_toggle_btn = tk.Button(self.toolbar_frame, command=self._toggle_settings_panel)
        self.settings_toggle_btn.pack(side=tk.RIGHT, padx=(8, 0))
        self.log_toggle_btn = tk.Button(self.toolbar_frame, command=self._toggle_log_panel)
        self.log_toggle_btn.pack(side=tk.RIGHT)

        self.settings_panel = tk.Frame(root)

        top = tk.Frame(self.settings_panel)
        top.pack(fill=tk.X, pady=(0, 10))

        self.edition_frame = tk.LabelFrame(top, padx=10, pady=8)
        self.edition_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.edition_standard_radio = tk.Radiobutton(
            self.edition_frame,
            variable=self.edition_var,
            value="standard",
            command=self._on_edition_change,
        )
        self.edition_standard_radio.pack(side=tk.LEFT, padx=(0, 12))

        self.edition_advanced_radio = tk.Radiobutton(
            self.edition_frame,
            variable=self.edition_var,
            value="advanced",
            command=self._on_edition_change,
        )
        self.edition_advanced_radio.pack(side=tk.LEFT)

        self.lang_frame = tk.LabelFrame(top, padx=10, pady=8)
        self.lang_frame.pack(side=tk.LEFT, padx=(10, 0))

        self.lang_zh_radio = tk.Radiobutton(
            self.lang_frame,
            text="中文",
            variable=self.lang_var,
            value="zh",
            command=self._on_language_change,
        )
        self.lang_zh_radio.pack(side=tk.LEFT, padx=(0, 8))

        self.lang_en_radio = tk.Radiobutton(
            self.lang_frame,
            text="English",
            variable=self.lang_var,
            value="en",
            command=self._on_language_change,
        )
        self.lang_en_radio.pack(side=tk.LEFT)

        self.standard_frame = tk.LabelFrame(root, padx=10, pady=10)
        self.advanced_frame = tk.LabelFrame(root, padx=10, pady=10)

        self._build_standard_frame()
        self._build_advanced_frame()

        self.log_panel = tk.LabelFrame(root, padx=10, pady=10)
        self.log_box = ScrolledText(self.log_panel, height=10, wrap=tk.WORD)
        self.log_box.pack(fill=tk.BOTH, expand=True)
        self.log_box.configure(state=tk.DISABLED)

        self.status_frame = tk.Frame(root, pady=8)
        self.status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_label = tk.Label(self.status_frame, anchor="w")
        self.status_label.pack(fill=tk.X)

        self._register_interactive_widgets()
        self._on_standard_mode_change()
        self._on_advanced_mode_change()
        self._on_edition_change()

    def _build_standard_frame(self) -> None:
        self.standard_mode_frame = tk.LabelFrame(self.standard_frame, padx=10, pady=8)
        self.standard_mode_frame.pack(fill=tk.X, pady=(0, 10))

        self.standard_paste_radio = tk.Radiobutton(
            self.standard_mode_frame,
            variable=self.standard_mode_var,
            value="paste",
            command=self._on_standard_mode_change,
        )
        self.standard_paste_radio.pack(side=tk.LEFT, padx=(0, 12))

        self.standard_single_radio = tk.Radiobutton(
            self.standard_mode_frame,
            variable=self.standard_mode_var,
            value="single",
            command=self._on_standard_mode_change,
        )
        self.standard_single_radio.pack(side=tk.LEFT)

        self.standard_paste_panel = tk.Frame(self.standard_frame)
        self.standard_paste_content_frame = tk.Frame(self.standard_paste_panel)
        self.standard_paste_content_frame.pack(fill=tk.X)

        self.std_paste_content_label = tk.Label(self.standard_paste_content_frame)
        self.std_paste_content_label.pack(anchor="w")
        self.std_paste_hint_label = tk.Label(self.standard_paste_content_frame, anchor="w", justify=tk.LEFT)
        self.std_paste_hint_label.pack(anchor="w", pady=(2, 6))

        self.std_paste_text = ScrolledText(self.standard_paste_content_frame, height=8, wrap=tk.WORD)
        self.std_paste_text.pack(fill=tk.X, pady=(4, 8))

        self.standard_paste_form_frame = tk.Frame(self.standard_paste_panel)
        self.standard_paste_form_frame.pack(fill=tk.X)

        self.std_paste_output_label = tk.Label(self.standard_paste_form_frame)
        self.std_paste_output_label.grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.std_paste_output_entry = tk.Entry(self.standard_paste_form_frame, textvariable=self.std_paste_output_var)
        self.std_paste_output_entry.grid(row=0, column=1, sticky="ew", pady=4)
        self.std_paste_output_btn = tk.Button(self.standard_paste_form_frame, command=self._pick_std_paste_output)
        self.std_paste_output_btn.grid(row=0, column=2, padx=(8, 0), pady=4)

        self.std_paste_filename_label = tk.Label(self.standard_paste_form_frame)
        self.std_paste_filename_label.grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.std_paste_filename_entry = tk.Entry(
            self.standard_paste_form_frame,
            textvariable=self.std_paste_filename_var,
        )
        self.std_paste_filename_entry.grid(row=1, column=1, sticky="ew", pady=4)

        self.std_paste_template_label = tk.Label(self.standard_paste_form_frame)
        self.std_paste_template_label.grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self.std_paste_template_entry = tk.Entry(self.standard_paste_form_frame, textvariable=self.std_paste_template_var)
        self.std_paste_template_entry.grid(row=2, column=1, sticky="ew", pady=4)
        self.std_paste_template_btn = tk.Button(
            self.standard_paste_form_frame,
            command=lambda: self._pick_template(self.std_paste_template_var),
        )
        self.std_paste_template_btn.grid(row=2, column=2, padx=(8, 0), pady=4)
        self.std_paste_template_clear_btn = tk.Button(
            self.standard_paste_form_frame,
            width=8,
            command=lambda: self.std_paste_template_var.set(""),
        )
        self.std_paste_template_clear_btn.grid(row=2, column=3, padx=(8, 0), pady=4)

        self.std_paste_convert_btn = tk.Button(
            self.standard_paste_form_frame,
            width=18,
            command=self._start_standard_paste_convert,
            bg="#1f6feb",
            fg="white",
        )
        self.std_paste_convert_btn.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self.standard_paste_form_frame.columnconfigure(1, weight=1)

        self.standard_single_panel = tk.Frame(self.standard_frame)

        self.std_single_input_label = tk.Label(self.standard_single_panel)
        self.std_single_input_label.grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.std_single_input_entry = tk.Entry(self.standard_single_panel, textvariable=self.std_single_input_var)
        self.std_single_input_entry.grid(row=0, column=1, sticky="ew", pady=4)
        self.std_single_input_btn = tk.Button(self.standard_single_panel, command=self._pick_std_single_input)
        self.std_single_input_btn.grid(row=0, column=2, padx=(8, 0), pady=4)

        self.std_single_output_label = tk.Label(self.standard_single_panel)
        self.std_single_output_label.grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.std_single_output_entry = tk.Entry(self.standard_single_panel, textvariable=self.std_single_output_var)
        self.std_single_output_entry.grid(row=1, column=1, sticky="ew", pady=4)
        self.std_single_output_btn = tk.Button(self.standard_single_panel, command=self._pick_std_single_output)
        self.std_single_output_btn.grid(row=1, column=2, padx=(8, 0), pady=4)

        self.std_single_template_label = tk.Label(self.standard_single_panel)
        self.std_single_template_label.grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self.std_single_template_entry = tk.Entry(self.standard_single_panel, textvariable=self.std_single_template_var)
        self.std_single_template_entry.grid(row=2, column=1, sticky="ew", pady=4)
        self.std_single_template_btn = tk.Button(
            self.standard_single_panel,
            command=lambda: self._pick_template(self.std_single_template_var),
        )
        self.std_single_template_btn.grid(row=2, column=2, padx=(8, 0), pady=4)
        self.std_single_template_clear_btn = tk.Button(
            self.standard_single_panel,
            width=8,
            command=lambda: self.std_single_template_var.set(""),
        )
        self.std_single_template_clear_btn.grid(row=2, column=3, padx=(8, 0), pady=4)

        self.std_single_convert_btn = tk.Button(
            self.standard_single_panel,
            width=18,
            command=self._start_standard_single_convert,
            bg="#1f6feb",
            fg="white",
        )
        self.std_single_convert_btn.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self.standard_single_panel.columnconfigure(1, weight=1)

    def _build_advanced_frame(self) -> None:
        self.advanced_mode_frame = tk.LabelFrame(self.advanced_frame, padx=10, pady=8)
        self.advanced_mode_frame.pack(fill=tk.X, pady=(0, 10))

        self.adv_single_radio = tk.Radiobutton(
            self.advanced_mode_frame,
            variable=self.advanced_mode_var,
            value="single",
            command=self._on_advanced_mode_change,
        )
        self.adv_single_radio.pack(side=tk.LEFT, padx=(0, 12))

        self.adv_folder_radio = tk.Radiobutton(
            self.advanced_mode_frame,
            variable=self.advanced_mode_var,
            value="folder",
            command=self._on_advanced_mode_change,
        )
        self.adv_folder_radio.pack(side=tk.LEFT)

        self.advanced_path_frame = tk.LabelFrame(self.advanced_frame, padx=10, pady=10)
        self.advanced_path_frame.pack(fill=tk.X, pady=(0, 10))

        self.adv_input_label = tk.Label(self.advanced_path_frame)
        self.adv_input_label.grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.adv_input_entry = tk.Entry(self.advanced_path_frame, textvariable=self.adv_input_var)
        self.adv_input_entry.grid(row=0, column=1, sticky="ew", pady=4)
        self.adv_input_btn = tk.Button(self.advanced_path_frame, command=self._pick_adv_input)
        self.adv_input_btn.grid(row=0, column=2, padx=(8, 0), pady=4)

        self.adv_output_label = tk.Label(self.advanced_path_frame)
        self.adv_output_label.grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.adv_output_entry = tk.Entry(self.advanced_path_frame, textvariable=self.adv_output_var)
        self.adv_output_entry.grid(row=1, column=1, sticky="ew", pady=4)
        self.adv_output_btn = tk.Button(self.advanced_path_frame, command=self._pick_adv_output)
        self.adv_output_btn.grid(row=1, column=2, padx=(8, 0), pady=4)

        self.adv_template_label = tk.Label(self.advanced_path_frame)
        self.adv_template_label.grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self.adv_template_entry = tk.Entry(self.advanced_path_frame, textvariable=self.adv_template_var)
        self.adv_template_entry.grid(row=2, column=1, sticky="ew", pady=4)
        self.adv_template_btn = tk.Button(
            self.advanced_path_frame,
            command=lambda: self._pick_template(self.adv_template_var),
        )
        self.adv_template_btn.grid(row=2, column=2, padx=(8, 0), pady=4)
        self.adv_template_clear_btn = tk.Button(
            self.advanced_path_frame,
            width=8,
            command=lambda: self.adv_template_var.set(""),
        )
        self.adv_template_clear_btn.grid(row=2, column=3, padx=(8, 0), pady=4)

        self.advanced_path_frame.columnconfigure(1, weight=1)

        self.adv_option_frame = tk.LabelFrame(self.advanced_frame, padx=10, pady=8)
        self.adv_option_frame.pack(fill=tk.X, pady=(0, 10))

        self.adv_recursive_check = tk.Checkbutton(self.adv_option_frame, variable=self.adv_recursive_var)
        self.adv_recursive_check.pack(anchor="w")

        self.adv_convert_btn = tk.Button(
            self.advanced_frame,
            width=18,
            command=self._start_advanced_convert,
            bg="#1f6feb",
            fg="white",
        )
        self.adv_convert_btn.pack(anchor="w")

    def _register_interactive_widgets(self) -> None:
        self.interactive_widgets = [
            self.settings_toggle_btn,
            self.log_toggle_btn,
            self.edition_standard_radio,
            self.edition_advanced_radio,
            self.lang_zh_radio,
            self.lang_en_radio,
            self.standard_paste_radio,
            self.standard_single_radio,
            self.adv_single_radio,
            self.adv_folder_radio,
            self.adv_recursive_check,
            self.std_paste_output_entry,
            self.std_paste_output_btn,
            self.std_paste_filename_entry,
            self.std_paste_template_entry,
            self.std_paste_template_btn,
            self.std_paste_template_clear_btn,
            self.std_paste_text,
            self.std_paste_convert_btn,
            self.std_single_input_entry,
            self.std_single_input_btn,
            self.std_single_output_entry,
            self.std_single_output_btn,
            self.std_single_template_entry,
            self.std_single_template_btn,
            self.std_single_template_clear_btn,
            self.std_single_convert_btn,
            self.adv_input_entry,
            self.adv_input_btn,
            self.adv_output_entry,
            self.adv_output_btn,
            self.adv_template_entry,
            self.adv_template_btn,
            self.adv_template_clear_btn,
            self.adv_convert_btn,
        ]

    def _apply_modern_theme(self) -> None:
        self._style_widget_tree(self)

        self.header_title_label.configure(
            font=self.fonts["title"],
            fg=self.colors["text"],
        )
        self.header_subtitle_label.configure(
            font=self.fonts["subtitle"],
            fg=self.colors["muted"],
        )

        self._style_secondary_button(self.settings_toggle_btn)
        self._style_secondary_button(self.log_toggle_btn)

        self._style_primary_button(self.std_paste_convert_btn)
        self._style_primary_button(self.std_single_convert_btn)
        self._style_primary_button(self.adv_convert_btn)

        for btn in [
            self.std_paste_output_btn,
            self.std_paste_template_btn,
            self.std_paste_template_clear_btn,
            self.std_single_input_btn,
            self.std_single_output_btn,
            self.std_single_template_btn,
            self.std_single_template_clear_btn,
            self.adv_input_btn,
            self.adv_output_btn,
            self.adv_template_btn,
            self.adv_template_clear_btn,
        ]:
            self._style_secondary_button(btn)

        self.std_paste_text.configure(
            bg=self.colors["card"],
            fg=self.colors["text"],
            insertbackground=self.colors["text"],
            relief="flat",
            padx=10,
            pady=8,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["primary"],
        )
        self.log_box.configure(
            bg=self.colors["log_bg"],
            fg=self.colors["log_fg"],
            insertbackground=self.colors["log_fg"],
            relief="flat",
            padx=10,
            pady=8,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["primary"],
            font=self.fonts["mono"],
        )
        self.status_frame.configure(bg=self.colors["card"])
        self.status_label.configure(
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=self.fonts["heading"],
            padx=10,
            pady=6,
        )

    def _style_widget_tree(self, widget: tk.Misc) -> None:
        self._style_one_widget(widget)
        for child in widget.winfo_children():
            self._style_widget_tree(child)

    def _style_one_widget(self, widget: tk.Misc) -> None:
        if isinstance(widget, tk.Tk):
            widget.configure(bg=self.colors["bg"])
            return

        parent_bg = self.colors["bg"]
        master = getattr(widget, "master", None)
        if master is not None:
            try:
                parent_bg = master.cget("bg")
            except tk.TclError:
                parent_bg = self.colors["bg"]

        if isinstance(widget, tk.LabelFrame):
            widget.configure(
                bg=self.colors["card"],
                fg=self.colors["text"],
                bd=1,
                relief="solid",
                font=self.fonts["heading"],
                highlightthickness=1,
                highlightbackground=self.colors["border"],
                highlightcolor=self.colors["border"],
                padx=12,
                pady=10,
            )
            return

        if isinstance(widget, tk.Frame):
            widget.configure(bg=parent_bg)
            return

        if isinstance(widget, tk.Label):
            widget.configure(bg=parent_bg, fg=self.colors["text"])
            return

        if isinstance(widget, tk.Entry):
            widget.configure(
                bg=self.colors["card"],
                fg=self.colors["text"],
                relief="flat",
                insertbackground=self.colors["text"],
                highlightthickness=1,
                highlightbackground=self.colors["border"],
                highlightcolor=self.colors["primary"],
            )
            return

        if isinstance(widget, tk.Radiobutton):
            widget.configure(
                bg=parent_bg,
                fg=self.colors["text"],
                activebackground=parent_bg,
                activeforeground=self.colors["text"],
                selectcolor=self.colors["card"],
                highlightthickness=0,
                bd=0,
            )
            return

        if isinstance(widget, tk.Checkbutton):
            widget.configure(
                bg=parent_bg,
                fg=self.colors["text"],
                activebackground=parent_bg,
                activeforeground=self.colors["text"],
                selectcolor=self.colors["card"],
                highlightthickness=0,
                bd=0,
            )
            return

        if isinstance(widget, tk.Button):
            self._style_secondary_button(widget)

    def _style_primary_button(self, btn: tk.Button) -> None:
        btn.configure(
            bg=self.colors["primary"],
            fg="white",
            activebackground=self.colors["primary_hover"],
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=16,
            pady=9,
            font=self.fonts["button"],
            cursor="hand2",
        )

    def _style_secondary_button(self, btn: tk.Button) -> None:
        btn.configure(
            bg=self.colors["button"],
            fg=self.colors["text"],
            activebackground=self.colors["button_hover"],
            activeforeground=self.colors["text"],
            relief="flat",
            bd=0,
            padx=12,
            pady=7,
            cursor="hand2",
        )

    def _setup_paste_placeholder(self) -> None:
        self.std_paste_text.bind("<FocusIn>", self._on_paste_focus_in)
        self.std_paste_text.bind("<FocusOut>", self._on_paste_focus_out)
        self._set_paste_placeholder_text()

    def _set_paste_placeholder_text(self) -> None:
        self.std_paste_text.configure(state=tk.NORMAL)
        self.std_paste_text.delete("1.0", tk.END)
        self.std_paste_text.insert("1.0", self.t("paste_placeholder"))
        self.std_paste_text.configure(fg=self.colors["muted"])
        self.paste_placeholder_active = True

    def _on_paste_focus_in(self, _event: tk.Event) -> None:
        if not self.paste_placeholder_active:
            return
        self.std_paste_text.delete("1.0", tk.END)
        self.std_paste_text.configure(fg=self.colors["text"])
        self.paste_placeholder_active = False

    def _on_paste_focus_out(self, _event: tk.Event) -> None:
        if self.std_paste_text.get("1.0", tk.END).strip():
            return
        self._set_paste_placeholder_text()

    def _get_paste_content(self) -> str:
        if self.paste_placeholder_active:
            return ""
        return self.std_paste_text.get("1.0", tk.END).strip()

    def _refresh_texts(self) -> None:
        self.title(self.t("window_title"))
        self.header_title_label.configure(text=self.t("header_title"))
        self.header_subtitle_label.configure(text=self.t("header_subtitle"))
        self.log_panel.configure(text=self.t("logs_title"))

        self.settings_toggle_btn.configure(
            text=self.t("settings_hide_btn") if self.settings_visible else self.t("settings_show_btn")
        )
        self.log_toggle_btn.configure(
            text=self.t("logs_hide_btn") if self.log_visible else self.t("logs_show_btn")
        )

        self.lang_frame.configure(text=self.t("language_group"))
        self.edition_frame.configure(text=self.t("edition_group"))
        self.edition_standard_radio.configure(text=self.t("edition_standard"))
        self.edition_advanced_radio.configure(text=self.t("edition_advanced"))

        self.standard_frame.configure(text=self.t("standard_title"))
        self.standard_mode_frame.configure(text=self.t("standard_mode_group"))
        self.standard_paste_radio.configure(text=self.t("standard_mode_paste"))
        self.standard_single_radio.configure(text=self.t("standard_mode_single"))

        self.std_paste_content_label.configure(text=self.t("markdown_content_label"))
        self.std_paste_hint_label.configure(text=self.t("paste_hint"))
        self.std_paste_output_label.configure(text=self.t("output_label"))
        self.std_paste_filename_label.configure(text=self.t("filename_label"))
        self.std_paste_template_label.configure(text=self.t("template_label"))
        self.std_paste_output_btn.configure(text=self.t("browse_btn"))
        self.std_paste_template_btn.configure(text=self.t("browse_btn"))
        self.std_paste_template_clear_btn.configure(text=self.t("clear_btn"))

        self.std_single_input_label.configure(text=self.t("input_label"))
        self.std_single_output_label.configure(text=self.t("output_label"))
        self.std_single_template_label.configure(text=self.t("template_label"))
        self.std_single_input_btn.configure(text=self.t("browse_btn"))
        self.std_single_output_btn.configure(text=self.t("browse_btn"))
        self.std_single_template_btn.configure(text=self.t("browse_btn"))
        self.std_single_template_clear_btn.configure(text=self.t("clear_btn"))

        self.advanced_frame.configure(text=self.t("advanced_title"))
        self.advanced_mode_frame.configure(text=self.t("advanced_mode_group"))
        self.adv_single_radio.configure(text=self.t("advanced_mode_single"))
        self.adv_folder_radio.configure(text=self.t("advanced_mode_folder"))
        self.advanced_path_frame.configure(text=self.t("advanced_title"))
        self.adv_input_label.configure(text=self.t("input_label"))
        self.adv_output_label.configure(text=self.t("output_label"))
        self.adv_template_label.configure(text=self.t("template_label"))
        self.adv_input_btn.configure(text=self.t("browse_btn"))
        self.adv_output_btn.configure(text=self.t("browse_btn"))
        self.adv_template_btn.configure(text=self.t("browse_btn"))
        self.adv_template_clear_btn.configure(text=self.t("clear_btn"))
        self.adv_option_frame.configure(text=self.t("advanced_title"))
        self.adv_recursive_check.configure(text=self.t("advanced_recursive"))

        convert_text = self.t("converting_btn") if self.is_running else self.t("convert_btn")
        self.std_paste_convert_btn.configure(text=convert_text)
        self.std_single_convert_btn.configure(text=convert_text)
        self.adv_convert_btn.configure(text=convert_text)

        if self.paste_placeholder_active and not self.is_running:
            self._set_paste_placeholder_text()

    def _on_language_change(self) -> None:
        self._refresh_texts()
        self._set_status(self.status_kind)

    def _toggle_settings_panel(self) -> None:
        if self.is_running:
            return

        self.settings_visible = not self.settings_visible
        if self.settings_visible:
            before_widget = self.standard_frame if self.edition_var.get() == "standard" else self.advanced_frame
            self.settings_panel.pack(fill=tk.X, pady=(0, 8), before=before_widget)
        else:
            self.settings_panel.pack_forget()
        self._refresh_texts()

    def _toggle_log_panel(self) -> None:
        self.log_visible = not self.log_visible
        if self.log_visible:
            self.log_panel.pack(fill=tk.BOTH, expand=True, pady=(0, 8), before=self.status_frame)
        else:
            self.log_panel.pack_forget()
        self._refresh_texts()

    def _set_status(self, kind: str) -> None:
        self.status_kind = kind
        if kind == "running":
            text = self.t("status_running")
            color = self.colors["primary"]
        elif kind == "success":
            text = self.t("status_success")
            color = "#15803d"
        elif kind == "partial":
            text = self.t("status_partial")
            color = "#b45309"
        elif kind == "fail":
            text = self.t("status_fail")
            color = "#b91c1c"
        elif kind == "empty":
            text = self.t("status_empty")
            color = self.colors["muted"]
        else:
            text = self.t("status_ready")
            color = self.colors["muted"]

        self.status_label.configure(text=text, fg=color)

    def _handle_result(self, ok: int, fail: int, lang: str) -> None:
        if ok > 0 and fail == 0:
            self._set_status("success")
            messagebox.showinfo(
                self.t("result_success_title", lang=lang),
                self.t("result_success_msg", lang=lang, ok=ok),
            )
            return

        if ok > 0 and fail > 0:
            self._set_status("partial")
            messagebox.showwarning(
                self.t("result_partial_title", lang=lang),
                self.t("result_partial_msg", lang=lang, ok=ok, fail=fail),
            )
            return

        if ok == 0 and fail == 0:
            self._set_status("empty")
            messagebox.showwarning(
                self.t("result_empty_title", lang=lang),
                self.t("result_empty_msg", lang=lang),
            )
            return

        self._set_status("fail")
        messagebox.showerror(
            self.t("result_fail_title", lang=lang),
            self.t("result_fail_msg", lang=lang, fail=fail),
        )

    def _on_edition_change(self) -> None:
        edition = self.edition_var.get()
        self.standard_frame.pack_forget()
        self.advanced_frame.pack_forget()
        before_widget: tk.Widget = self.log_panel if self.log_visible else self.status_frame

        if edition == "standard":
            self.standard_frame.pack(fill=tk.X, pady=(0, 10), before=before_widget)
        else:
            self.advanced_frame.pack(fill=tk.X, pady=(0, 10), before=before_widget)

    def _on_standard_mode_change(self) -> None:
        mode = self.standard_mode_var.get()
        self.standard_paste_panel.pack_forget()
        self.standard_single_panel.pack_forget()

        if mode == "paste":
            self.standard_paste_panel.pack(fill=tk.X)
        else:
            self.standard_single_panel.pack(fill=tk.X)

    def _on_advanced_mode_change(self) -> None:
        folder_mode = self.advanced_mode_var.get() == "folder"
        if self.is_running:
            state = tk.DISABLED
        else:
            state = tk.NORMAL if folder_mode else tk.DISABLED
        self.adv_recursive_check.configure(state=state)

    def _apply_startup_args(self) -> None:
        if len(sys.argv) < 2:
            return

        arg_path = Path(sys.argv[1]).expanduser()
        if not arg_path.exists():
            return

        if arg_path.is_file():
            self.edition_var.set("standard")
            self.standard_mode_var.set("single")
            self.std_single_input_var.set(str(arg_path))
            if not self.std_single_output_var.get():
                self.std_single_output_var.set(str(arg_path.parent / "docx_output"))
        elif arg_path.is_dir():
            self.edition_var.set("advanced")
            self.advanced_mode_var.set("folder")
            self.adv_input_var.set(str(arg_path))
            if not self.adv_output_var.get():
                self.adv_output_var.set(str(arg_path / "docx_output"))

        self._on_standard_mode_change()
        self._on_advanced_mode_change()
        self._on_edition_change()

    def _pick_template(self, target_var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(
            title=self.t("choose_template"),
            filetypes=[("Word docx", "*.docx")],
        )
        if path:
            target_var.set(path)

    def _pick_std_paste_output(self) -> None:
        path = filedialog.askdirectory(title=self.t("choose_output_folder"))
        if path:
            self.std_paste_output_var.set(path)

    def _pick_std_single_input(self) -> None:
        path = filedialog.askopenfilename(
            title=self.t("choose_md_file"),
            filetypes=[("Markdown", "*.md *.markdown *.mdown *.mkd"), ("All files", "*.*")],
        )
        if path:
            self.std_single_input_var.set(path)
            if not self.std_single_output_var.get():
                path_obj = Path(path)
                self.std_single_output_var.set(str(path_obj.parent / "docx_output"))

    def _pick_std_single_output(self) -> None:
        path = filedialog.askdirectory(title=self.t("choose_output_folder"))
        if path:
            self.std_single_output_var.set(path)

    def _pick_adv_input(self) -> None:
        if self.advanced_mode_var.get() == "single":
            path = filedialog.askopenfilename(
                title=self.t("choose_md_file"),
                filetypes=[("Markdown", "*.md *.markdown *.mdown *.mkd"), ("All files", "*.*")],
            )
        else:
            path = filedialog.askdirectory(title=self.t("choose_source_folder"))

        if path:
            self.adv_input_var.set(path)
            if not self.adv_output_var.get():
                path_obj = Path(path)
                if self.advanced_mode_var.get() == "single":
                    self.adv_output_var.set(str(path_obj.parent / "docx_output"))
                else:
                    self.adv_output_var.set(str(path_obj / "docx_output"))

    def _pick_adv_output(self) -> None:
        path = filedialog.askdirectory(title=self.t("choose_output_folder"))
        if path:
            self.adv_output_var.set(path)

    def _resolve_pandoc_path(self) -> str | None:
        pandoc_name = "pandoc.exe" if sys.platform.startswith("win") else "pandoc"
        candidates: list[Path] = []

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

    def _validate_template(self, template: str) -> bool:
        if not template:
            return True

        template_path = Path(template)
        if template_path.is_file() and template_path.suffix.lower() == ".docx":
            return True

        messagebox.showerror(self.t("invalid_template_title"), self.t("invalid_template_msg"))
        return False

    def _normalize_docx_name(self, name: str) -> str | None:
        base = name.strip()
        if base.lower().endswith(".docx"):
            base = base[:-5]
        base = base.strip()

        if not base:
            return None

        if any(ch in INVALID_FILENAME_CHARS for ch in base):
            return None

        return f"{base}.docx"

    def _set_running(self, running: bool) -> None:
        self.is_running = running
        state = tk.DISABLED if running else tk.NORMAL

        for widget in self.interactive_widgets:
            try:
                widget.configure(state=state)
            except tk.TclError:
                pass

        self._on_advanced_mode_change()
        self._refresh_texts()
        if not running and not self.std_paste_text.get("1.0", tk.END).strip():
            self._set_paste_placeholder_text()

    def _start_standard_paste_convert(self) -> None:
        if self.is_running:
            return

        pandoc_path = self._resolve_pandoc_path()
        if not pandoc_path:
            messagebox.showerror(self.t("pandoc_not_found_title"), self.t("pandoc_not_found_msg"))
            return

        content = self._get_paste_content()
        output_dir = self.std_paste_output_var.get().strip()
        filename = self.std_paste_filename_var.get().strip()
        template = self.std_paste_template_var.get().strip()

        if not content:
            messagebox.showwarning(self.t("missing_content_title"), self.t("missing_content_msg"))
            return

        if not output_dir:
            messagebox.showwarning(self.t("missing_output_title"), self.t("missing_output_msg"))
            return

        if not filename:
            messagebox.showwarning(self.t("missing_filename_title"), self.t("missing_filename_msg"))
            return

        docx_name = self._normalize_docx_name(filename)
        if not docx_name:
            messagebox.showerror(self.t("invalid_filename_title"), self.t("invalid_filename_msg"))
            return

        if not self._validate_template(template):
            return

        out_dir = Path(output_dir)
        dst = out_dir / docx_name

        self._set_running(True)
        self._set_status("running")
        lang = self.lang_var.get()
        self._append_log("-" * 72)
        self._append_log(self.t("log_using_pandoc", lang=lang, pandoc=pandoc_path))
        self._append_log(self.t("log_mode", lang=lang, mode=self.t("mode_standard_paste", lang=lang)))
        self._append_log(self.t("log_output", lang=lang, output=out_dir))
        self._append_log(self.t("log_template", lang=lang, template=(template if template else self.t("none_text", lang=lang))))
        self._append_log(self.t("log_target", lang=lang, target=dst))

        thread = threading.Thread(
            target=self._convert_from_text_worker,
            args=(content, dst, template, pandoc_path, lang),
            daemon=True,
        )
        thread.start()

    def _start_standard_single_convert(self) -> None:
        if self.is_running:
            return

        pandoc_path = self._resolve_pandoc_path()
        if not pandoc_path:
            messagebox.showerror(self.t("pandoc_not_found_title"), self.t("pandoc_not_found_msg"))
            return

        input_path = self.std_single_input_var.get().strip()
        output_dir = self.std_single_output_var.get().strip()
        template = self.std_single_template_var.get().strip()

        if not input_path:
            messagebox.showwarning(self.t("missing_input_title"), self.t("missing_input_msg"))
            return

        if not output_dir:
            messagebox.showwarning(self.t("missing_output_title"), self.t("missing_output_msg"))
            return

        src = Path(input_path)
        if not src.is_file() or src.suffix.lower() not in MD_SUFFIXES:
            messagebox.showerror(self.t("invalid_input_title"), self.t("invalid_input_file_msg"))
            return

        if not self._validate_template(template):
            return

        out_dir = Path(output_dir)
        dst = out_dir / f"{src.stem}.docx"

        self._set_running(True)
        self._set_status("running")
        lang = self.lang_var.get()
        self._append_log("-" * 72)
        self._append_log(self.t("log_using_pandoc", lang=lang, pandoc=pandoc_path))
        self._append_log(self.t("log_mode", lang=lang, mode=self.t("mode_standard_single", lang=lang)))
        self._append_log(self.t("log_input", lang=lang, input=src))
        self._append_log(self.t("log_output", lang=lang, output=out_dir))
        self._append_log(self.t("log_template", lang=lang, template=(template if template else self.t("none_text", lang=lang))))

        thread = threading.Thread(
            target=self._convert_single_file_worker,
            args=(src, dst, template, pandoc_path, lang),
            daemon=True,
        )
        thread.start()

    def _start_advanced_convert(self) -> None:
        if self.is_running:
            return

        pandoc_path = self._resolve_pandoc_path()
        if not pandoc_path:
            messagebox.showerror(self.t("pandoc_not_found_title"), self.t("pandoc_not_found_msg"))
            return

        mode = self.advanced_mode_var.get()
        input_path = self.adv_input_var.get().strip()
        output_dir = self.adv_output_var.get().strip()
        template = self.adv_template_var.get().strip()
        recursive = self.adv_recursive_var.get()

        if not input_path:
            messagebox.showwarning(self.t("missing_input_title"), self.t("missing_input_msg"))
            return

        if not output_dir:
            messagebox.showwarning(self.t("missing_output_title"), self.t("missing_output_msg"))
            return

        src_path = Path(input_path)
        if mode == "single":
            if not src_path.is_file() or src_path.suffix.lower() not in MD_SUFFIXES:
                messagebox.showerror(self.t("invalid_input_title"), self.t("invalid_input_file_msg"))
                return
        else:
            if not src_path.is_dir():
                messagebox.showerror(self.t("invalid_input_title"), self.t("invalid_input_folder_msg"))
                return

        if not self._validate_template(template):
            return

        out_dir = Path(output_dir)

        self._set_running(True)
        self._set_status("running")
        lang = self.lang_var.get()
        mode_label = self.t("mode_advanced_single", lang=lang) if mode == "single" else self.t("mode_advanced_folder", lang=lang)
        self._append_log("-" * 72)
        self._append_log(self.t("log_using_pandoc", lang=lang, pandoc=pandoc_path))
        self._append_log(self.t("log_mode", lang=lang, mode=mode_label))
        self._append_log(self.t("log_input", lang=lang, input=src_path))
        self._append_log(self.t("log_output", lang=lang, output=out_dir))
        self._append_log(self.t("log_template", lang=lang, template=(template if template else self.t("none_text", lang=lang))))

        thread = threading.Thread(
            target=self._convert_advanced_worker,
            args=(src_path, out_dir, template, mode, recursive, pandoc_path, lang),
            daemon=True,
        )
        thread.start()

    def _convert_from_text_worker(
        self,
        content: str,
        dst: Path,
        template: str,
        pandoc_path: str,
        lang: str,
    ) -> None:
        ok = 0
        fail = 0

        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            try:
                self._run_pandoc_text(content, dst, template, pandoc_path)
                ok = 1
                self.log_queue.put(self.t("log_ok", lang=lang))
            except Exception as exc:
                fail = 1
                self.log_queue.put(self.t("log_fail", lang=lang, err=str(exc)))

            self.log_queue.put("=" * 72)
            self.log_queue.put(self.t("log_done", lang=lang, ok=ok, fail=fail))
            self.log_queue.put({"type": "result", "ok": ok, "fail": fail, "lang": lang})
        finally:
            self.log_queue.put("__UI_UNLOCK__")

    def _convert_single_file_worker(
        self,
        src: Path,
        dst: Path,
        template: str,
        pandoc_path: str,
        lang: str,
    ) -> None:
        ok = 0
        fail = 0

        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            self.log_queue.put(f"[1/1] {src} -> {dst}")
            try:
                self._run_pandoc_file(src, dst, template, pandoc_path)
                ok = 1
                self.log_queue.put(self.t("log_ok", lang=lang))
            except Exception as exc:
                fail = 1
                self.log_queue.put(self.t("log_fail", lang=lang, err=str(exc)))

            self.log_queue.put("=" * 72)
            self.log_queue.put(self.t("log_done", lang=lang, ok=ok, fail=fail))
            self.log_queue.put({"type": "result", "ok": ok, "fail": fail, "lang": lang})
        finally:
            self.log_queue.put("__UI_UNLOCK__")

    def _convert_advanced_worker(
        self,
        src_path: Path,
        out_dir: Path,
        template: str,
        mode: str,
        recursive: bool,
        pandoc_path: str,
        lang: str,
    ) -> None:
        ok = 0
        fail = 0

        try:
            out_dir.mkdir(parents=True, exist_ok=True)
            md_files = self._collect_md_files(src_path, mode, recursive)

            if not md_files:
                self.log_queue.put(self.t("log_no_md", lang=lang))
                self.log_queue.put("=" * 72)
                self.log_queue.put(self.t("log_done", lang=lang, ok=0, fail=0))
                self.log_queue.put({"type": "result", "ok": 0, "fail": 0, "lang": lang})
                return

            self.log_queue.put(self.t("log_found_files", lang=lang, count=len(md_files)))

            for idx, src in enumerate(md_files, start=1):
                if mode == "single":
                    dst = out_dir / f"{src.stem}.docx"
                else:
                    rel = src.relative_to(src_path)
                    dst = out_dir / rel.with_suffix(".docx")

                dst.parent.mkdir(parents=True, exist_ok=True)
                self.log_queue.put(f"[{idx}/{len(md_files)}] {src} -> {dst}")

                try:
                    self._run_pandoc_file(src, dst, template, pandoc_path)
                    ok += 1
                    self.log_queue.put(self.t("log_ok", lang=lang))
                except Exception as exc:
                    fail += 1
                    self.log_queue.put(self.t("log_fail", lang=lang, err=str(exc)))

            self.log_queue.put("=" * 72)
            self.log_queue.put(self.t("log_done", lang=lang, ok=ok, fail=fail))
            self.log_queue.put({"type": "result", "ok": ok, "fail": fail, "lang": lang})
        finally:
            self.log_queue.put("__UI_UNLOCK__")

    def _collect_md_files(self, in_path: Path, mode: str, recursive: bool) -> list[Path]:
        if mode == "single":
            return [in_path]

        pattern = "**/*" if recursive else "*"
        files = sorted(in_path.glob(pattern))
        return [f for f in files if f.is_file() and f.suffix.lower() in MD_SUFFIXES]

    def _run_pandoc_file(self, src: Path, dst: Path, template: str, pandoc_path: str) -> None:
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

    def _run_pandoc_text(self, content: str, dst: Path, template: str, pandoc_path: str) -> None:
        cmd = [
            pandoc_path,
            "-f",
            "markdown",
            "-t",
            "docx",
            "-o",
            str(dst),
        ]

        if template:
            cmd.extend(["--reference-doc", template])

        result = subprocess.run(cmd, input=content.encode("utf-8"), capture_output=True)
        if result.returncode != 0:
            stderr = result.stderr.decode("utf-8", errors="replace").strip() if result.stderr else ""
            stdout = result.stdout.decode("utf-8", errors="replace").strip() if result.stdout else ""
            err = stderr or stdout or "pandoc conversion failed"
            raise RuntimeError(err)

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
                elif isinstance(msg, dict) and msg.get("type") == "result":
                    self._handle_result(
                        int(msg.get("ok", 0)),
                        int(msg.get("fail", 0)),
                        str(msg.get("lang", self.lang_var.get())),
                    )
                else:
                    self._append_log(msg)
        except queue.Empty:
            pass
        finally:
            self.after(120, self._drain_log_queue)


if __name__ == "__main__":
    app = Md2DocxApp()
    app.mainloop()
