"""
doc2md GUI - 文档转 Markdown 可视化工具
支持拖拽文件/文件夹、选择文件、实时进度、转换日志、中英文切换

作者: 倪刚
"""

import os
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

__version__ = "1.2.0"
__author__ = "倪刚"

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))
from doc2md import convert_docx, convert_doc, convert_pdf, PDF_ENGINES, PDF_ENGINE_LABELS

SUPPORTED_EXTS = {".doc", ".docx", ".pdf"}

DND_AVAILABLE = False
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    pass

try:
    import customtkinter as ctk
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    CTK_AVAILABLE = True
except ImportError:
    CTK_AVAILABLE = False


# ─── 多语言 ──────────────────────────────────────────────────────────────────

LANG = {
    "zh": {
        "title": "doc2md",
        "subtitle": "doc / docx / pdf  ->  Markdown + 图片  (供 AI 阅读)",
        "author_info": f"v{__version__}  作者: {__author__}",
        "drop_hint_dnd": "拖拽文件或文件夹到此处，或使用下方按钮",
        "drop_hint_btn": "使用下方按钮添加文件",
        "btn_files": "选择文件",
        "btn_folder": "选择文件夹",
        "btn_clear": "清空",
        "output_label": "输出目录:",
        "output_placeholder": "(与源文件同目录)",
        "embed_check": "图片内嵌 (base64)",
        "engine_label": "PDF引擎:",
        "btn_convert": "开始转换",
        "btn_converting": "转换中...",
        "btn_copy_log": "复制日志",
        "btn_open_dir": "打开目录",
        "status_ready": "就绪",
        "log_ready": "doc2md 就绪。",
        "log_dnd_on": "拖拽功能已启用 - 可直接拖入文件或文件夹！",
        "log_dnd_off": "拖拽不可用，请使用「选择文件」/「选择文件夹」按钮。",
        "log_no_doc": "[!] 拖入的内容中未找到 doc/docx/pdf 文件",
        "log_cleared": "已清空所有文件。",
        "log_added": "+ 添加了 {added} 个文件，共 {total} 个",
        "files_ready": "已选择 {n} 个文件",
        "dlg_select_files": "选择文档",
        "dlg_select_folder": "选择文件夹",
        "dlg_select_output": "选择输出目录",
        "dlg_empty": "未找到文档",
        "dlg_empty_msg": "该文件夹中没有 doc/docx/pdf 文件:\n{folder}",
        "dlg_no_files": "未添加文件",
        "dlg_no_files_msg": "请先添加文件！",
        "log_converting": "--- [{idx}/{total}] {name} ---",
        "status_converting": "[{idx}/{total}] {name}",
        "status_step": "[{idx}/{total}] {name} - {step}",
        "log_ok": "  完成 -> {path} ({elapsed})",
        "log_fail": "  失败: {err}",
        "done_summary": "完成！成功 {ok}，失败 {fail}，共 {total}，耗时 {elapsed}",
        "log_copied": "日志已复制到剪贴板",
        "lang_label": "语言",
    },
    "en": {
        "title": "doc2md",
        "subtitle": "doc / docx / pdf  ->  Markdown + Images  (for AI)",
        "author_info": f"v{__version__}  Author: Ni Gang",
        "drop_hint_dnd": "Drag files/folders here, or use buttons below",
        "drop_hint_btn": "Use buttons below to add files",
        "btn_files": "Select Files",
        "btn_folder": "Select Folder",
        "btn_clear": "Clear All",
        "output_label": "Output dir:",
        "output_placeholder": "(same as source file)",
        "embed_check": "Embed images (base64)",
        "engine_label": "PDF Engine:",
        "btn_convert": "Start Convert",
        "btn_converting": "Converting...",
        "btn_copy_log": "Copy Log",
        "btn_open_dir": "Open Dir",
        "status_ready": "Ready",
        "log_ready": "doc2md ready.",
        "log_dnd_on": "Drag & drop enabled - drop files or folders here!",
        "log_dnd_off": "Drag & drop not available, use 'Select Files' / 'Select Folder' buttons.",
        "log_no_doc": "[!] No doc/docx/pdf found in dropped items",
        "log_cleared": "All files cleared.",
        "log_added": "+ {added} file(s) added, total: {total}",
        "files_ready": "{n} file(s) ready",
        "dlg_select_files": "Select documents",
        "dlg_select_folder": "Select folder",
        "dlg_select_output": "Select output directory",
        "dlg_empty": "Empty",
        "dlg_empty_msg": "No doc/docx/pdf found in:\n{folder}",
        "dlg_no_files": "No files",
        "dlg_no_files_msg": "Please add files first!",
        "log_converting": "--- [{idx}/{total}] {name} ---",
        "status_converting": "[{idx}/{total}] {name}",
        "status_step": "[{idx}/{total}] {name} - {step}",
        "log_ok": "  OK -> {path} ({elapsed})",
        "log_fail": "  FAILED: {err}",
        "done_summary": "Done! {ok} ok, {fail} failed, {total} total, took {elapsed}",
        "log_copied": "Log copied to clipboard",
        "lang_label": "Lang",
    },
}


def _create_root():
    if DND_AVAILABLE and CTK_AVAILABLE:
        class DnDCTk(ctk.CTk, TkinterDnD.DnDWrapper):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.TkdndVersion = TkinterDnD._require(self)
        return DnDCTk()
    if CTK_AVAILABLE:
        return ctk.CTk()
    if DND_AVAILABLE:
        return TkinterDnD.Tk()
    return tk.Tk()


def _parse_drop_data(raw: str) -> list[str]:
    files = []
    in_brace = False
    current = ""
    for ch in raw:
        if ch == "{":
            in_brace = True
            continue
        if ch == "}":
            in_brace = False
            if current:
                files.append(current)
            current = ""
            continue
        if ch == " " and not in_brace:
            if current:
                files.append(current)
            current = ""
            continue
        current += ch
    if current:
        files.append(current)
    return files


class Doc2MdApp:
    def __init__(self):
        self.root = _create_root()
        self.root.title("doc2md")
        self.root.geometry("760x640")
        self.root.minsize(600, 500)

        self.file_list: list[Path] = []
        self.converting = False
        self.cur_lang = "zh"
        self._widgets = {}

        self._setup_dnd()
        self._build_ui()

    def t(self, key: str, **kwargs) -> str:
        text = LANG[self.cur_lang].get(key, key)
        if kwargs:
            text = text.format(**kwargs)
        return text

    def _setup_dnd(self):
        self.dnd_ok = False
        if DND_AVAILABLE:
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self._on_drop)
                self.dnd_ok = True
            except Exception:
                pass

    def _on_drop(self, event):
        paths = _parse_drop_data(event.data)
        all_files = []
        for p_str in paths:
            p = Path(p_str)
            if p.is_dir():
                found = list(p.glob("*.doc")) + list(p.glob("*.docx")) + list(p.glob("*.pdf"))
                all_files.extend(found)
            elif p.suffix.lower() in SUPPORTED_EXTS:
                all_files.append(p)
        if all_files:
            self._add_files(all_files)
        else:
            self._log(self.t("log_no_doc"))

    # ─── UI ───────────────────────────────────────────────────────────

    def _build_ui(self):
        root = self.root
        root.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(3, weight=1)

        # --- Title row (with lang switcher on the right) ---
        if CTK_AVAILABLE:
            title_frame = ctk.CTkFrame(root, fg_color="transparent")
        else:
            title_frame = tk.Frame(root)
        title_frame.grid(row=0, column=0, padx=20, pady=(15, 5), sticky="ew")
        title_frame.grid_columnconfigure(0, weight=1)

        if CTK_AVAILABLE:
            left = ctk.CTkFrame(title_frame, fg_color="transparent")
            left.grid(row=0, column=0, sticky="w")
            self._widgets["title"] = ctk.CTkLabel(left, text=self.t("title"),
                                                    font=ctk.CTkFont(size=26, weight="bold"))
            self._widgets["title"].pack(anchor="w")
            self._widgets["subtitle"] = ctk.CTkLabel(left, text=self.t("subtitle"),
                                                       font=ctk.CTkFont(size=12), text_color="gray55")
            self._widgets["subtitle"].pack(anchor="w")

            right = ctk.CTkFrame(title_frame, fg_color="transparent")
            right.grid(row=0, column=1, sticky="ne", padx=(10, 0))
            self.lang_var = ctk.StringVar(value="中文")
            self._widgets["lang_menu"] = ctk.CTkOptionMenu(
                right, values=["中文", "English"],
                variable=self.lang_var, width=100,
                command=self._on_lang_change,
                font=ctk.CTkFont(size=12),
            )
            self._widgets["lang_menu"].pack(anchor="e")
            self._widgets["author_label"] = ctk.CTkLabel(
                right, text=self.t("author_info"),
                font=ctk.CTkFont(size=10), text_color="gray45")
            self._widgets["author_label"].pack(anchor="e", pady=(4, 0))
        else:
            left = tk.Frame(title_frame)
            left.grid(row=0, column=0, sticky="w")
            self._widgets["title"] = tk.Label(left, text=self.t("title"), font=("", 20, "bold"))
            self._widgets["title"].pack(anchor="w")
            self._widgets["subtitle"] = tk.Label(left, text=self.t("subtitle"), fg="gray")
            self._widgets["subtitle"].pack(anchor="w")

            right = tk.Frame(title_frame)
            right.grid(row=0, column=1, sticky="ne")
            self.lang_var = tk.StringVar(value="中文")
            self._widgets["lang_menu"] = tk.OptionMenu(
                right, self.lang_var, "中文", "English",
                command=self._on_lang_change)
            self._widgets["lang_menu"].pack(anchor="e")
            self._widgets["author_label"] = tk.Label(
                right, text=self.t("author_info"), fg="gray", font=("", 8))
            self._widgets["author_label"].pack(anchor="e")

        # --- Drop zone ---
        if CTK_AVAILABLE:
            drop_frame = ctk.CTkFrame(root, corner_radius=12, border_width=2,
                                       border_color=("gray70", "gray30"))
        else:
            drop_frame = tk.LabelFrame(root, padx=10, pady=10)
        drop_frame.grid(row=1, column=0, padx=20, pady=8, sticky="ew")
        drop_frame.grid_columnconfigure(0, weight=1)

        hint = self.t("drop_hint_dnd") if self.dnd_ok else self.t("drop_hint_btn")
        if CTK_AVAILABLE:
            self.drop_label = ctk.CTkLabel(drop_frame, text=hint,
                                            font=ctk.CTkFont(size=13), text_color="gray50")
        else:
            self.drop_label = tk.Label(drop_frame, text=hint, fg="gray")
        self.drop_label.grid(row=0, column=0, pady=(12, 8), columnspan=4)

        btn_cfg = {"row": 1, "pady": (0, 12)}
        if CTK_AVAILABLE:
            self._widgets["btn_files"] = ctk.CTkButton(drop_frame, text=self.t("btn_files"), width=140,
                                                         command=self._pick_files)
            self._widgets["btn_files"].grid(column=0, padx=6, **btn_cfg)
            self._widgets["btn_folder"] = ctk.CTkButton(drop_frame, text=self.t("btn_folder"), width=140,
                                                          command=self._pick_folder)
            self._widgets["btn_folder"].grid(column=1, padx=6, **btn_cfg)
            self._widgets["btn_clear"] = ctk.CTkButton(drop_frame, text=self.t("btn_clear"), width=100,
                                                         fg_color=("gray70", "gray30"),
                                                         hover_color=("gray60", "gray40"),
                                                         command=self._clear_files)
            self._widgets["btn_clear"].grid(column=2, padx=6, **btn_cfg)
        else:
            self._widgets["btn_files"] = tk.Button(drop_frame, text=self.t("btn_files"), width=16,
                                                     command=self._pick_files)
            self._widgets["btn_files"].grid(column=0, padx=4, **btn_cfg)
            self._widgets["btn_folder"] = tk.Button(drop_frame, text=self.t("btn_folder"), width=16,
                                                      command=self._pick_folder)
            self._widgets["btn_folder"].grid(column=1, padx=4, **btn_cfg)
            self._widgets["btn_clear"] = tk.Button(drop_frame, text=self.t("btn_clear"), width=12,
                                                     command=self._clear_files)
            self._widgets["btn_clear"].grid(column=2, padx=4, **btn_cfg)

        # --- Options row ---
        if CTK_AVAILABLE:
            opts = ctk.CTkFrame(root, fg_color="transparent")
        else:
            opts = tk.Frame(root)
        opts.grid(row=2, column=0, padx=20, pady=4, sticky="ew")
        opts.grid_columnconfigure(1, weight=1)

        if CTK_AVAILABLE:
            self._widgets["output_label"] = ctk.CTkLabel(opts, text=self.t("output_label"),
                                                           font=ctk.CTkFont(size=12))
            self._widgets["output_label"].grid(row=0, column=0, padx=(0, 6))
            self.output_entry = ctk.CTkEntry(opts, placeholder_text=self.t("output_placeholder"))
            self.output_entry.grid(row=0, column=1, sticky="ew")
            ctk.CTkButton(opts, text="...", width=36, command=self._pick_output).grid(row=0, column=2, padx=(4, 0))
            self.embed_var = ctk.BooleanVar(value=False)
            self._widgets["embed_check"] = ctk.CTkCheckBox(opts, text=self.t("embed_check"),
                                                             variable=self.embed_var,
                                                             font=ctk.CTkFont(size=11))
            self._widgets["embed_check"].grid(row=0, column=3, padx=(12, 0))

            self._widgets["engine_label"] = ctk.CTkLabel(opts, text=self.t("engine_label"),
                                                           font=ctk.CTkFont(size=12))
            self._widgets["engine_label"].grid(row=1, column=0, padx=(0, 6), pady=(6, 0))
            self.engine_var = ctk.StringVar(value="auto")
            engine_values = [f"{e} - {PDF_ENGINE_LABELS[e]}" for e in PDF_ENGINES]
            self._widgets["engine_combo"] = ctk.CTkComboBox(opts, values=engine_values,
                                                              variable=self.engine_var,
                                                              font=ctk.CTkFont(size=11),
                                                              state="readonly", width=360)
            self._widgets["engine_combo"].set(engine_values[0])
            self._widgets["engine_combo"].grid(row=1, column=1, columnspan=3, sticky="ew", pady=(6, 0))
        else:
            self._widgets["output_label"] = tk.Label(opts, text=self.t("output_label"))
            self._widgets["output_label"].grid(row=0, column=0, padx=(0, 4))
            self.output_entry = tk.Entry(opts)
            self.output_entry.grid(row=0, column=1, sticky="ew")
            tk.Button(opts, text="...", width=3, command=self._pick_output).grid(row=0, column=2, padx=(4, 0))
            self.embed_var = tk.BooleanVar(value=False)
            self._widgets["embed_check"] = tk.Checkbutton(opts, text=self.t("embed_check"),
                                                            variable=self.embed_var)
            self._widgets["embed_check"].grid(row=0, column=3, padx=(8, 0))

            tk.Label(opts, text=self.t("engine_label")).grid(row=1, column=0, padx=(0, 4), pady=(4, 0))
            self.engine_var = tk.StringVar(value="auto")
            engine_values = [f"{e} - {PDF_ENGINE_LABELS[e]}" for e in PDF_ENGINES]
            import tkinter.ttk as ttk
            self._widgets["engine_combo"] = ttk.Combobox(opts, textvariable=self.engine_var,
                                                           values=engine_values, state="readonly")
            self._widgets["engine_combo"].set(engine_values[0])
            self._widgets["engine_combo"].grid(row=1, column=1, columnspan=3, sticky="ew", pady=(4, 0))

        # --- Log area ---
        if CTK_AVAILABLE:
            log_frame = ctk.CTkFrame(root, corner_radius=10)
            log_frame.grid(row=3, column=0, padx=20, pady=6, sticky="nsew")
            log_frame.grid_columnconfigure(0, weight=1)
            log_frame.grid_rowconfigure(0, weight=1)
            self.log_text = ctk.CTkTextbox(log_frame, font=ctk.CTkFont(family="Consolas", size=12),
                                            state="disabled", wrap="word")
            self.log_text.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        else:
            log_frame = tk.Frame(root)
            log_frame.grid(row=3, column=0, padx=20, pady=6, sticky="nsew")
            log_frame.grid_columnconfigure(0, weight=1)
            log_frame.grid_rowconfigure(0, weight=1)
            self.log_text = tk.Text(log_frame, font=("Consolas", 10), state="disabled", wrap="word")
            self.log_text.grid(row=0, column=0, sticky="nsew")
            sb = tk.Scrollbar(log_frame, command=self.log_text.yview)
            sb.grid(row=0, column=1, sticky="ns")
            self.log_text.config(yscrollcommand=sb.set)

        # --- Log toolbar ---
        if CTK_AVAILABLE:
            log_toolbar = ctk.CTkFrame(root, fg_color="transparent")
        else:
            log_toolbar = tk.Frame(root)
        log_toolbar.grid(row=4, column=0, padx=20, pady=(0, 2), sticky="ew")

        if CTK_AVAILABLE:
            self._widgets["btn_copy_log"] = ctk.CTkButton(
                log_toolbar, text=self.t("btn_copy_log"), width=90, height=26,
                font=ctk.CTkFont(size=11),
                fg_color=("gray70", "gray30"), hover_color=("gray60", "gray40"),
                command=self._copy_log)
            self._widgets["btn_copy_log"].pack(side="right", padx=(4, 0))
            self._widgets["btn_open_dir"] = ctk.CTkButton(
                log_toolbar, text=self.t("btn_open_dir"), width=90, height=26,
                font=ctk.CTkFont(size=11),
                fg_color=("gray70", "gray30"), hover_color=("gray60", "gray40"),
                command=self._open_output_dir)
            self._widgets["btn_open_dir"].pack(side="right", padx=(4, 0))
        else:
            self._widgets["btn_copy_log"] = tk.Button(
                log_toolbar, text=self.t("btn_copy_log"), width=10,
                command=self._copy_log)
            self._widgets["btn_copy_log"].pack(side="right", padx=(4, 0))
            self._widgets["btn_open_dir"] = tk.Button(
                log_toolbar, text=self.t("btn_open_dir"), width=10,
                command=self._open_output_dir)
            self._widgets["btn_open_dir"].pack(side="right", padx=(4, 0))

        # --- Bottom bar ---
        if CTK_AVAILABLE:
            bottom = ctk.CTkFrame(root, fg_color="transparent")
        else:
            bottom = tk.Frame(root)
        bottom.grid(row=5, column=0, padx=20, pady=(4, 15), sticky="ew")
        bottom.grid_columnconfigure(0, weight=1)

        if CTK_AVAILABLE:
            self.progress = ctk.CTkProgressBar(bottom, height=8)
            self.progress.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
            self.progress.set(0)
            self._widgets["status"] = ctk.CTkLabel(bottom, text=self.t("status_ready"),
                                                     font=ctk.CTkFont(size=12), text_color="gray50")
            self._widgets["status"].grid(row=1, column=0, sticky="w")
            self.status_label = self._widgets["status"]
            self._widgets["btn_convert"] = ctk.CTkButton(bottom, text=self.t("btn_convert"),
                                                           width=160, height=38,
                                                           font=ctk.CTkFont(size=14, weight="bold"),
                                                           command=self._start_convert)
            self._widgets["btn_convert"].grid(row=1, column=1, sticky="e")
            self.convert_btn = self._widgets["btn_convert"]
        else:
            self._widgets["status"] = tk.Label(bottom, text=self.t("status_ready"), fg="gray")
            self._widgets["status"].grid(row=0, column=0, sticky="w")
            self.status_label = self._widgets["status"]
            self._widgets["btn_convert"] = tk.Button(bottom, text=self.t("btn_convert"), width=20,
                                                       command=self._start_convert)
            self._widgets["btn_convert"].grid(row=0, column=1, sticky="e")
            self.convert_btn = self._widgets["btn_convert"]
            self.progress = None

        self.last_output_dir = None

        self._log(self.t("log_ready"))
        if self.dnd_ok:
            self._log(self.t("log_dnd_on"))
        else:
            self._log(self.t("log_dnd_off"))

    # ─── Language switch ──────────────────────────────────────────────

    def _on_lang_change(self, choice):
        self.cur_lang = "en" if choice == "English" else "zh"
        self._refresh_texts()

    def _refresh_texts(self):
        _cfg = "configure" if CTK_AVAILABLE else "config"

        getattr(self._widgets["title"], _cfg)(text=self.t("title"))
        getattr(self._widgets["subtitle"], _cfg)(text=self.t("subtitle"))
        getattr(self._widgets["author_label"], _cfg)(text=self.t("author_info"))
        getattr(self._widgets["btn_files"], _cfg)(text=self.t("btn_files"))
        getattr(self._widgets["btn_folder"], _cfg)(text=self.t("btn_folder"))
        getattr(self._widgets["btn_clear"], _cfg)(text=self.t("btn_clear"))
        getattr(self._widgets["output_label"], _cfg)(text=self.t("output_label"))
        getattr(self._widgets["btn_copy_log"], _cfg)(text=self.t("btn_copy_log"))
        getattr(self._widgets["btn_open_dir"], _cfg)(text=self.t("btn_open_dir"))
        getattr(self._widgets["status"], _cfg)(text=self.t("status_ready"))

        if CTK_AVAILABLE:
            self._widgets["embed_check"].configure(text=self.t("embed_check"))
            self._widgets["engine_label"].configure(text=self.t("engine_label"))
            self.output_entry.configure(placeholder_text=self.t("output_placeholder"))
        else:
            self._widgets["embed_check"].config(text=self.t("embed_check"))

        if not self.converting:
            getattr(self._widgets["btn_convert"], _cfg)(text=self.t("btn_convert"))

        if not self.file_list:
            hint = self.t("drop_hint_dnd") if self.dnd_ok else self.t("drop_hint_btn")
            getattr(self.drop_label, _cfg)(text=hint)
        else:
            getattr(self.drop_label, _cfg)(text=self.t("files_ready", n=len(self.file_list)))

        self._log(f"[{self.cur_lang.upper()}] " + self.t("log_ready"))

    # ─── Actions ──────────────────────────────────────────────────────

    def _log(self, msg: str):
        if CTK_AVAILABLE:
            self.log_text.configure(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        else:
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")

    def _set_status(self, text: str):
        _cfg = "configure" if CTK_AVAILABLE else "config"
        getattr(self.status_label, _cfg)(text=text)

    def _set_progress(self, val: float):
        if self.progress:
            self.progress.set(val)

    def _pick_files(self):
        files = filedialog.askopenfilenames(
            title=self.t("dlg_select_files"),
            filetypes=[("Documents", "*.doc *.docx *.pdf"), ("All", "*.*")],
        )
        if files:
            self._add_files([Path(f) for f in files])

    def _pick_folder(self):
        folder = filedialog.askdirectory(title=self.t("dlg_select_folder"))
        if folder:
            p = Path(folder)
            found = list(p.glob("*.doc")) + list(p.glob("*.docx")) + list(p.glob("*.pdf"))
            if found:
                self._add_files(found)
            else:
                messagebox.showinfo(self.t("dlg_empty"),
                                    self.t("dlg_empty_msg", folder=folder))

    def _pick_output(self):
        folder = filedialog.askdirectory(title=self.t("dlg_select_output"))
        if folder:
            if CTK_AVAILABLE:
                self.output_entry.delete(0, "end")
                self.output_entry.insert(0, folder)
            else:
                self.output_entry.delete(0, tk.END)
                self.output_entry.insert(0, folder)

    def _add_files(self, files: list[Path]):
        existing = {str(f) for f in self.file_list}
        added = 0
        new_files = []
        for f in files:
            key = str(f.resolve())
            if key not in existing and f.suffix.lower() in SUPPORTED_EXTS:
                self.file_list.append(f)
                existing.add(key)
                new_files.append(f)
                added += 1

        _cfg = "configure" if CTK_AVAILABLE else "config"
        getattr(self.drop_label, _cfg)(text=self.t("files_ready", n=len(self.file_list)))

        if added:
            self._log(f"\n{self.t('log_added', added=added, total=len(self.file_list))}")
            for f in new_files:
                size_kb = f.stat().st_size / 1024
                ext = f.suffix.lower()
                tag = {".doc": "DOC", ".docx": "DOCX", ".pdf": "PDF"}.get(ext, "?")
                if size_kb >= 1024:
                    size_str = f"{size_kb / 1024:.1f} MB"
                else:
                    size_str = f"{size_kb:.0f} KB"
                self._log(f"  [{tag}] {f.name}  ({size_str})")

    def _clear_files(self):
        self.file_list.clear()
        hint = self.t("drop_hint_dnd") if self.dnd_ok else self.t("drop_hint_btn")
        _cfg = "configure" if CTK_AVAILABLE else "config"
        getattr(self.drop_label, _cfg)(text=hint)
        self._log(self.t("log_cleared"))

    def _start_convert(self):
        if self.converting:
            return
        if not self.file_list:
            messagebox.showwarning(self.t("dlg_no_files"), self.t("dlg_no_files_msg"))
            return

        self.converting = True
        _cfg = "configure" if CTK_AVAILABLE else "config"
        getattr(self.convert_btn, _cfg)(state="disabled", text=self.t("btn_converting"))
        self._set_progress(0)

        thread = threading.Thread(target=self._do_convert, daemon=True)
        thread.start()

    def _do_convert(self):
        output_text = self.output_entry.get().strip()
        embed = self.embed_var.get()
        total = len(self.file_list)
        success = 0
        failed = 0
        total_start = time.time()
        last_out_dir = None

        for idx, fpath in enumerate(self.file_list):
            base_pct = idx / total
            i = idx + 1
            self.root.after(0, self._set_status,
                            self.t("status_converting", idx=i, total=total, name=fpath.name))
            self.root.after(0, self._log,
                            f"\n{self.t('log_converting', idx=i, total=total, name=fpath.name)}")

            out_dir = Path(output_text) if output_text else fpath.parent
            out_dir.mkdir(parents=True, exist_ok=True)
            last_out_dir = out_dir

            current_idx = idx

            def progress_cb(pct, step, _idx=current_idx, _name=fpath.name):
                overall = base_pct + (pct / 100) / total
                self.root.after(0, self._set_progress, overall)
                self.root.after(0, self._set_status,
                                self.t("status_step", idx=_idx + 1, total=total, name=_name, step=step))

            converters = {".docx": convert_docx, ".doc": convert_doc, ".pdf": convert_pdf}
            converter = converters.get(fpath.suffix.lower())

            if converter:
                file_start = time.time()
                try:
                    if converter == convert_pdf:
                        engine = self.engine_var.get().split(" - ")[0].strip()
                        md_path = converter(fpath, out_dir, embed, progress_cb, engine=engine)
                    else:
                        md_path = converter(fpath, out_dir, embed, progress_cb)
                    elapsed = _fmt_elapsed(time.time() - file_start)
                    self.root.after(0, self._log,
                                    self.t("log_ok", path=md_path, elapsed=elapsed))
                    success += 1
                except Exception as e:
                    self.root.after(0, self._log, self.t("log_fail", err=e))
                    failed += 1

        self.last_output_dir = last_out_dir
        total_elapsed = _fmt_elapsed(time.time() - total_start)
        self.root.after(0, self._set_progress, 1.0)
        summary = self.t("done_summary", ok=success, fail=failed,
                         total=total, elapsed=total_elapsed)
        self.root.after(0, self._set_status, summary)
        self.root.after(0, self._log, f"\n{'=' * 40}\n{summary}\n")
        self.root.after(0, self._finish_convert)

    def _finish_convert(self):
        self.converting = False
        _cfg = "configure" if CTK_AVAILABLE else "config"
        getattr(self.convert_btn, _cfg)(state="normal", text=self.t("btn_convert"))
        try:
            self.root.bell()
        except Exception:
            pass

    def _copy_log(self):
        if CTK_AVAILABLE:
            self.log_text.configure(state="normal")
            content = self.log_text.get("1.0", "end").strip()
            self.log_text.configure(state="disabled")
        else:
            self.log_text.config(state="normal")
            content = self.log_text.get("1.0", "end").strip()
            self.log_text.config(state="disabled")
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self._log(self.t("log_copied"))

    def _open_output_dir(self):
        target = None
        if self.last_output_dir and self.last_output_dir.exists():
            target = self.last_output_dir
        else:
            output_text = self.output_entry.get().strip()
            if output_text:
                p = Path(output_text)
                if p.exists():
                    target = p
            if not target and self.file_list:
                target = self.file_list[0].parent
        if target:
            if sys.platform == "win32":
                os.startfile(str(target))
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", str(target)])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", str(target)])

    def run(self):
        self.root.mainloop()


def _fmt_elapsed(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:.1f}s"
    m = int(seconds) // 60
    s = seconds - m * 60
    return f"{m}m{s:.0f}s"


def main():
    app = Doc2MdApp()
    app.run()


if __name__ == "__main__":
    main()
