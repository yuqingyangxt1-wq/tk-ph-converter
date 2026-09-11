"""Tkinter GUI for the TK Philippines table converter.

Layout (single-page, flat) — based on the v1.2 reference screenshot:

  ┌────────────────────────────────────────────────┐
  │ [ 转化 ]  库存池  设置                          │  ← top tab bar
  ├────────────────────────────────────────────────┤
  │ (body — content for the selected section)      │
  │                                                │
  │   转化 section:                                 │
  │     ┌──────────────────────────────────┐       │
  │     │  选择 xlsx  (drag/click area)    │       │
  │     │  源文件: 产品 X 个 · SKU Y 行    │       │
  │     │  文件/输出路径                  │       │
  │     │                                  │       │
  │     │  转化设置 (勾选才生效，会自动保存)│       │
  │     │   ☑ 标题前缀    [____]            │       │
  │     │   ☑ brand       [____]            │       │
  │     │   ☑ price       [____]            │       │
  │     │   ☑ quantity    [____]            │       │
  │     │   ☑ cod         [____]            │       │
  │     │   ☑ 标题随机后缀 [☐数量][☐长度]  │       │
  │     │   ☑ 拆分文件 (每份一个 xlsx)     │       │
  │     │   ☑ size chart  [____]            │       │
  │     │                                  │       │
  │     │  [开始转化]                       │       │
  │                                                │
  │   库存池 section:                              │
  │     池名 + 入池按钮 + 池列表                  │
  │                                                │
  │   设置 section:                                │
  │     全量字段设置 + 列映射                      │
  ├────────────────────────────────────────────────┤
  │  日志                                          │
  │  ┌──────────────────────────────────────────┐  │
  │  │   ScrolledText widget (append-only)       │  │
  │  └──────────────────────────────────────────┘  │
  │ 状态文本 · 版本号                                 │
  └────────────────────────────────────────────────┘

All long-running work runs on a background thread; UI updates are
posted back via a thread-safe queue so the interface never freezes.
"""
from __future__ import annotations

import os
import queue
import threading
import traceback
from pathlib import Path
from typing import Any, Callable

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from . import config as cfg_mod
from .config import (
    APP_NAME,
    default_config,
    get_app_dir,
    get_settings,
    load_config,
    save_config,
)
from .converter import convert_source
from .product_pool import (
    PoolInfo,
    add_to_pool,
    list_pools,
    load_pool,
)
from .source_reader import detect_format, read_source
from .tiktok_writer import template_path


# ---------------------------------------------------------------------------
# Worker bridge (background thread → UI queue)
# ---------------------------------------------------------------------------


class Worker:
    """Run a callable on a background thread and stream progress to the UI."""

    def __init__(self, app: "App"):
        self.app = app
        self.queue: queue.Queue = queue.Queue()
        self._t: threading.Thread | None = None

    def submit(self, fn, *args, on_done=None, **kwargs):
        if self._t and self._t.is_alive():
            messagebox.showwarning("提示", "有任务正在执行，请先等待完成。")
            return
        self.app.set_busy(True)

        def runner():
            try:
                def progress(p: float, text: str):
                    self.queue.put(("progress", p, text))
                def log(text: str):
                    self.queue.put(("log", text))
                result = fn(*args, progress=progress, log=log, **kwargs)
                self.queue.put(("done", result, on_done))
            except Exception as e:  # noqa: BLE001
                tb = traceback.format_exc()
                self.queue.put(("error", e, tb, on_done))

        self._t = threading.Thread(target=runner, daemon=True)
        self._t.start()
        self.app.after(80, self._drain)

    def _drain(self):
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]
                if kind == "progress":
                    _, p, text = item
                    self.app.on_progress(p, text)
                elif kind == "log":
                    _, text = item
                    self.app.append_log(text)
                elif kind == "done":
                    _, result, on_done = item
                    self.app.set_busy(False)
                    if on_done:
                        on_done(result)
                elif kind == "error":
                    _, exc, tb, on_done = item
                    self.app.set_busy(False)
                    self.app.on_progress(0, f"错误: {exc}")
                    self.app.append_log(f"[ERROR] {exc}\n{tb}")
                    messagebox.showerror("运行出错", f"{exc}\n\n{tb}")
                    if on_done:
                        on_done(None)
        except queue.Empty:
            pass
        if self._t and self._t.is_alive():
            self.app.after(80, self._drain)


# ---------------------------------------------------------------------------
# Tab bar (custom segmented control)
# ---------------------------------------------------------------------------


class TabBar(tk.Frame):
    """A flat, segmented-control style tab bar."""

    def __init__(self, master: tk.Misc, tabs: list[tuple[str, str]],
                 on_change: Callable[[str], None], **kwargs):
        super().__init__(master, **kwargs)
        self._on_change = on_change
        self._tabs: dict[str, tk.Button] = {}
        self._active: str | None = None
        for i, (key, label) in enumerate(tabs):
            b = tk.Button(
                self, text=label, bd=0, padx=18, pady=8,
                font=("", 10, "bold"),
                background="#f4f5f7",
                activebackground="#e8eaee",
                foreground="#222",
                activeforeground="#000",
                cursor="hand2",
            )
            b.bind("<Button-1>", lambda _e, k=key: self.select(k))
            b.grid(row=0, column=i, sticky="nsew")
            self._tabs[key] = b
        for k, _ in tabs:
            self.columnconfigure(list(self._tabs).index(k), weight=0)
        # Bottom border line
        sep = tk.Frame(self, height=1, bg="#d0d4dc")
        sep.grid(row=1, column=0, columnspan=len(tabs), sticky="ew")

    def select(self, key: str):
        for k, b in self._tabs.items():
            if k == key:
                b.configure(background="#ffffff", foreground="#0a58ff", relief="flat")
            else:
                b.configure(background="#f4f5f7", foreground="#222", relief="flat")
        self._active = key
        if self._on_change:
            self._on_change(key)


# ---------------------------------------------------------------------------
# Drag-or-click zone for the source xlsx
# ---------------------------------------------------------------------------


class DropZone(tk.Frame):
    """A big rectangular zone that responds to click (picks a file) and
    also tries to receive drop events from tkinterdnd2 when available.

    If tkinterdnd2 is unavailable the drop side is silently disabled —
    clicking still works.
    """

    def __init__(self, master: tk.Misc, on_file: Callable[[str], None], **kwargs):
        super().__init__(
            master,
            bg="#f7f8fb", highlightbackground="#c8ccd6",
            highlightthickness=2, bd=0,
            **kwargs,
        )
        self._on_file = on_file
        # Centered prompt text.
        self._lbl = tk.Label(
            self,
            text="拖入 xlsx 到此处，或下方选择文件",
            bg="#f7f8fb", fg="#666", font=("", 11),
            pady=22,
        )
        self._lbl.pack(expand=True, fill="both")
        self._lbl.bind("<Button-1>", lambda _e: self._pick())
        self.bind("<Button-1>", lambda _e: self._pick())
        # Best-effort: register as a drop target.
        try:
            from tkinterdnd2 import DND_FILES  # type: ignore
            self.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
            self.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore[attr-defined]
            self._lbl.configure(text="拖入 xlsx 到此处，或点击选择文件")
        except Exception:
            # Drag-and-drop disabled; click still works.
            pass

    def _pick(self):
        p = filedialog.askopenfilename(
            title="选择源表 xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if p:
            self._on_file(p)

    def _on_drop(self, event):
        # Tk dnd gives us a string like '{path1} {path2}' possibly with braces.
        data = event.data.strip()
        if data.startswith("{") and data.endswith("}"):
            data = data[1:-1]
        # Take the first whitespace-separated token.
        path = data.split(" ", 1)[0] if " " in data else data
        if path and os.path.isfile(path) and path.lower().endswith(".xlsx"):
            self._on_file(path)


# ---------------------------------------------------------------------------
# Compact "转化设置" panel
# ---------------------------------------------------------------------------


class ConvertSettingsPanel(tk.Frame):
    """A compact settings grid mirroring the v1.2 layout:

      ☑ label    [ editable field, optionally with sub-controls ]
           hint line in grey (when enabled)
    """

    def __init__(self, master: tk.Misc, settings: dict[str, Any],
                 on_change: Callable[[], None], **kwargs):
        super().__init__(master, bg="#ffffff", **kwargs)
        self._settings = settings
        self._on_change = on_change
        self._vars: dict[str, tk.Variable] = {}
        self._checks: dict[str, tk.BooleanVar] = {}
        self._build()

    # -- helpers ----------------------------------------------------------

    def _build(self):
        # Outer container with vertical padding.
        body = tk.Frame(self, bg="#ffffff")
        body.pack(fill="x", padx=20, pady=(8, 4))

        # Title row
        tk.Label(
            body, text="转化设置  (勾选才生效，会自动保存)",
            bg="#ffffff", fg="#222", font=("", 10, "bold"),
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        # Each setting row. We store
        #   (key_value, key_enabled_or_None, label, hint, kind)
        # so the auto-save layer can map back to the canonical settings dict.
        # kind ∈ {"text", "int"} — selects the widget type for value rows.
        # enabled_key=None means "no checkbox; this row is always-active".
        row_defs: list[tuple[str, str | None, str, str, str]] = [
            ("title_prefix",       "title_prefix_enabled",          "标题前缀",       "启用后去随机后缀并截断至 245 字符", "text"),
            ("brand_value",        "brand_enabled",                 "brand",          "启用后写入到 brand 列", "text"),
            ("price_value",        "price_enabled",                 "price",          "启用后覆盖每行 price", "int"),
            ("quantity_value",     "quantity_enabled",              "quantity",       "启用后覆盖每行 quantity", "int"),
            ("cod_value",          "cod_enabled",                   "cod",            "启用后写入到 cod 列（Y/N）", "text"),
            ("title_random",       "title_random_suffix_enabled",   "标题随机后缀",    "启用 + 截断至 N 位小写后缀，同一本尺码禁用相同后缀", "text"),
            ("split_output_files", "split_output_files",            "拆分文件",        "勾选后，每份输出一个独立 xlsx", "text"),
            ("output_copies",      None,                            "每产品输出份数",  "≥2 时启用防查重；<2 即只生成 1 个", "int"),
            ("size_chart_value",   "size_chart_enabled",            "size chart",     "留空用内置默认尺码图 URL", "text"),
        ]

        # We keep direct handles to each entry's ttk.Entry widget so that
        # future state-toggle code doesn't have to traverse the tree.
        self._entries: dict[str, ttk.Entry] = {}

        for i, (key, enabled_key, label, hint, kind) in enumerate(row_defs):
            r = i + 1

            if enabled_key is None:
                # Value-only row (no enable/disable checkbox).
                placeholder = tk.Label(  # noqa: F841 — keep grid column width
                    body, text=" ", bg="#ffffff", bd=0,
                )
                placeholder.grid(row=r, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
                ck = placeholder
            else:
                # Checkbox
                ck_var = tk.BooleanVar(value=bool(self._settings.get(enabled_key, False)))
                self._checks[enabled_key] = ck_var
                ck = tk.Checkbutton(
                    body, variable=ck_var, bg="#ffffff", bd=0, highlightthickness=0,
                    activebackground="#ffffff",
                    command=lambda k=enabled_key: self._on_check_change(k),
                )
                ck.grid(row=r, column=0, sticky="w", padx=(0, 8), pady=(6, 0))

            # Label
            lbl = tk.Label(
                body, text=label, bg="#ffffff", fg="#111",
                font=("Consolas", 10), width=18, anchor="w",
            )
            lbl.grid(row=r, column=1, sticky="w", pady=(6, 0))

            # Field
            val = self._settings.get(key, "")
            if isinstance(val, (int, float)):
                val = str(val)
            v = tk.StringVar(value=str(val))
            self._vars[key] = v
            if kind == "int":
                entry = ttk.Spinbox(
                    body, textvariable=v, from_=1, to=99, width=10,
                )
            else:
                entry = ttk.Entry(body, textvariable=v, width=40)
            entry.grid(row=r, column=2, sticky="ew", pady=(6, 0))
            self._entries[key] = entry
            v.trace_add("write", lambda *_, k=key: self._on_text_change(k))

            # Hint row (under the field, only shown when enabled)
            hint_lbl = tk.Label(
                body, text=hint, bg="#ffffff", fg="#999", font=("", 9), anchor="w",
            )
            hint_lbl.grid(row=r, column=2, sticky="w", padx=(0, 0), pady=(34, 0))

            body.columnconfigure(2, weight=1)
            if enabled_key is not None:
                self._apply_visual_state(enabled_key, self._checks[enabled_key].get())

        # Run button row
        btn_row = tk.Frame(self, bg="#ffffff")
        btn_row.pack(fill="x", padx=20, pady=(6, 12))
        self._run_btn = tk.Button(
            btn_row, text="开始转化", padx=24, pady=8,
            bg="#0a58ff", fg="white", activebackground="#0a58ff",
            activeforeground="white", bd=0, font=("", 10, "bold"),
            cursor="hand2",
        )
        self._run_btn.pack(side="left")
        # v3.2.4: 下载图片开关，默认勾上（8 线程并发）；想秒出表就取消
        self._download_imgs_var = tk.BooleanVar(
            value=bool(self._settings.get("download_images_enabled", True))
        )
        tk.Checkbutton(
            btn_row,
            text="下载产品图片（8 线程并发；不勾直接出表更快）",
            variable=self._download_imgs_var,
            bg="#ffffff", bd=0, highlightthickness=0,
            activebackground="#ffffff",
            command=self._on_download_toggle,
        ).pack(side="left", padx=(16, 0))

    def _on_download_toggle(self):
        self._settings["download_images_enabled"] = bool(self._download_imgs_var.get())
        if hasattr(self, "_on_change") and self._on_change:
            self._on_change()

    def is_download_images_enabled(self) -> bool:
        return bool(self._download_imgs_var.get())

    # -- state ------------------------------------------------------------

    def get_run_callback(self) -> Callable[[], None]:
        return self._run_btn.invoke

    def run_button(self) -> tk.Button:
        return self._run_btn

    def collect(self) -> dict[str, Any]:
        """Return updated settings to be merged back into config."""
        # Return both the enabled-flag changes AND the value rows (always
        # include value-only rows like output_copies — those have no
        # checkbox but the user can still edit them inline).
        out: dict[str, Any] = {}
        for enabled_key, var in self._checks.items():
            out[enabled_key] = bool(var.get())
        for key, var in self._vars.items():
            val = var.get()
            # Coerce numeric settings. output_copies lives in the compact
            # panel as a value-only row, so it's always written through.
            if key in ("price_value", "quantity_value", "random_suffix_length", "output_copies"):
                try:
                    val = max(1, int(float(val)))
                except (TypeError, ValueError):
                    val = 1
            out[key] = val
        return out

    def _on_check_change(self, key: str):
        enabled = bool(self._checks[key].get())
        self._apply_visual_state(key, enabled)
        self._on_change()

    def _on_text_change(self, _key: str):
        self._on_change()

    def _apply_visual_state(self, enabled_key: str, enabled: bool):
        # Map the *_enabled flag to its corresponding value key, then toggle
        # the matching Entry's disabled state. Cached handles mean we don't
        # have to traverse the widget tree.
        rev = {
            "title_prefix_enabled":        "title_prefix",
            "brand_enabled":               "brand_value",
            "price_enabled":               "price_value",
            "quantity_enabled":            "quantity_value",
            "cod_enabled":                 "cod_value",
            "title_random_suffix_enabled": "title_random",
            "split_output_files":          "split_output_files",
            "size_chart_enabled":          "size_chart_value",
        }
        key = rev.get(enabled_key)
        if not key:
            return
        entry = self._entries.get(key)
        if entry is None:
            return
        try:
            if enabled:
                entry.state(["!disabled"])
            else:
                entry.state(["disabled"])
        except tk.TclError:
            pass


# ---------------------------------------------------------------------------
# Main application window
# ---------------------------------------------------------------------------


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{cfg_mod.__version__}")
        self.geometry("1080x820")
        self.minsize(960, 720)
        self.configure(bg="#ffffff")

        self.cfg = load_config()
        self.worker = Worker(self)

        # Top-level section state
        self._section_keys: list[str] = ["convert", "pool", "settings"]
        self._sections: dict[str, tk.Frame] = {}
        self._build_chrome()

        # Cached source metadata (for the "产品 X 个 · SKU Y 行" line)
        self._source_meta: dict[str, Any] = {}

    # ----- UI skeleton -----------------------------------------------------

    def _build_chrome(self):
        # Tab bar at the top
        self._tabbar = TabBar(
            self,
            tabs=[
                ("convert",  " 转化 "),
                ("pool",     " 库存池 "),
                ("settings", " 设置 "),
            ],
            on_change=self._show_section,
        )
        self._tabbar.pack(side="top", fill="x")

        # Section container (single visible frame at a time)
        body_wrap = tk.Frame(self, bg="#ffffff")
        body_wrap.pack(side="top", fill="both", expand=True)

        self._section_container = body_wrap
        for k in self._section_keys:
            f = tk.Frame(body_wrap, bg="#ffffff")
            self._sections[k] = f

        # Bottom: log panel + status bar
        self._build_log_panel(self)

        # Build all sections
        self._build_section_convert(self._sections["convert"])
        self._build_section_pool(self._sections["pool"])
        self._build_section_settings(self._sections["settings"])

        # Initial visible tab
        self._tabbar.select("convert")

    def _show_section(self, key: str):
        for k, f in self._sections.items():
            if k == key:
                f.pack(fill="both", expand=True)
            else:
                f.pack_forget()

    # ----- Convert section ------------------------------------------------

    def _build_section_convert(self, parent: tk.Frame):
        # 1. Drag-or-click zone
        self._dropzone = DropZone(parent, on_file=self._on_source_dropped)
        self._dropzone.pack(fill="x", padx=20, pady=(16, 6))

        # 2. Source summary line
        self._var_source_summary = tk.StringVar(value="源文件: 尚未选择")
        summary = tk.Label(
            parent, textvariable=self._var_source_summary,
            bg="#ffffff", fg="#0a6", font=("", 10, "bold"),
        )
        summary.pack(anchor="w", padx=20, pady=(0, 6))

        # 3. Path rows: file + output
        path_frame = tk.Frame(parent, bg="#ffffff")
        path_frame.pack(fill="x", padx=20, pady=(0, 4))

        # File row
        tk.Label(path_frame, text="文件:", bg="#ffffff", fg="#111",
                 width=8, anchor="w").grid(row=0, column=0, sticky="w")
        self.var_source = tk.StringVar(value=self.cfg.get("product_xlsx_last", ""))
        ent_src = ttk.Entry(path_frame, textvariable=self.var_source)
        ent_src.grid(row=0, column=1, sticky="ew", padx=(4, 6))
        ttk.Button(path_frame, text="选择…", width=10,
                   command=self._pick_source).grid(row=0, column=2)

        # Output row
        tk.Label(path_frame, text="输出:", bg="#ffffff", fg="#111",
                 width=8, anchor="w").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.var_output = tk.StringVar(value=self.cfg.get("product_xlsx_output_dir", ""))
        ent_out = ttk.Entry(path_frame, textvariable=self.var_output)
        ent_out.grid(row=1, column=1, sticky="ew", padx=(4, 6), pady=(4, 0))
        ttk.Button(path_frame, text="选择…", width=10,
                   command=self._pick_output).grid(row=1, column=2, pady=(4, 0))
        ttk.Button(path_frame, text="清空", width=10,
                   command=lambda: self.var_output.set("")).grid(row=1, column=3, padx=(6, 0), pady=(4, 0))

        path_frame.columnconfigure(1, weight=1)

        # Helper text under path rows
        tk.Label(parent, text="留空 = 默认输出到 xlsx 同目录的 ready/",
                 bg="#ffffff", fg="#888", font=("", 9),
                 anchor="w").pack(anchor="w", padx=20, pady=(2, 8))

        # 4. Compact settings panel
        settings = get_settings(self.cfg)
        self._convert_panel = ConvertSettingsPanel(
            parent, settings=settings, on_change=self._auto_save_settings,
        )
        self._convert_panel.pack(fill="x", padx=12, pady=(0, 8))
        self._convert_panel.run_button().configure(command=self._do_convert)

        # Save the auto-save callback as a method for the panel to call.
        self._convert_panel._on_change = self._auto_save_settings  # noqa: SLF001

    # ----- Pool section ---------------------------------------------------

    def _build_section_pool(self, parent: tk.Frame):
        wrap = tk.Frame(parent, bg="#ffffff")
        wrap.pack(fill="both", expand=True, padx=20, pady=20)

        info = (
            "把当前源表里的产品加入指定的产品池。\n"
            "同一个池可以多次入池，工具会自动追加。"
        )
        tk.Label(wrap, text=info, bg="#ffffff", justify="left",
                 fg="#222").pack(anchor="w", pady=(0, 6))

        row = tk.Frame(wrap, bg="#ffffff"); row.pack(fill="x", pady=4)
        tk.Label(row, text="池名称：", bg="#ffffff").pack(side="left")
        self.var_pool_name = tk.StringVar(value="默认池")
        ttk.Entry(row, textvariable=self.var_pool_name, width=24).pack(side="left", padx=4)
        ttk.Button(row, text="加入产品池",
                   command=self._do_add_pool).pack(side="left", padx=8)
        ttk.Button(row, text="刷新",
                   command=self._refresh_pools).pack(side="left")

        ttk.Separator(wrap, orient="horizontal").pack(fill="x", pady=12)
        tk.Label(wrap, text="当前产品池中的池：", bg="#ffffff").pack(anchor="w")

        cols = ("name", "count", "updated")
        self.tree_pools = ttk.Treeview(wrap, columns=cols, show="headings", height=10)
        for c in cols:
            self.tree_pools.heading(c, text={"name": "池名称", "count": "行数",
                                              "updated": "更新时间"}[c])
            self.tree_pools.column(c, width=180 if c == "name" else 100, anchor="w")
        self.tree_pools.pack(fill="both", expand=True, pady=(4, 12))

        # Run-extract-and-convert button
        ttk.Button(wrap, text="提取并转化 →",
                   command=self._do_extract_convert).pack(anchor="e", pady=(0, 8))

    # ----- Settings section -----------------------------------------------

    def _build_section_settings(self, parent: tk.Frame):
        wrap = tk.Frame(parent, bg="#ffffff")
        wrap.pack(fill="both", expand=True, padx=20, pady=20)

        info = (
            "完整字段设置与源表列名映射。\n"
            "日常改动直接改顶部「转化设置」即可，这里用来调高级 / 长尾参数。"
        )
        tk.Label(wrap, text=info, bg="#ffffff", justify="left",
                 fg="#222").pack(anchor="w", pady=(0, 8))

        canvas = tk.Canvas(wrap, borderwidth=0, highlightthickness=0, bg="#ffffff")
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas, bg="#ffffff")
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(inner_id, width=e.width))

        s = get_settings(self.cfg)
        self.set_vars = {}
        self.set_checks = {}

        def add_section(title):
            tk.Label(inner, text=title, bg="#ffffff", fg="#222",
                     font=("", 11, "bold")).pack(anchor="w", pady=(12, 4))

        def add_str(key, label, multi=False):
            row = tk.Frame(inner, bg="#ffffff"); row.pack(fill="x", pady=2)
            tk.Label(row, text=label, bg="#ffffff", width=22, anchor="w").pack(side="left")
            if multi:
                txt = tk.Text(row, height=6, width=60, wrap="word")
                txt.insert("1.0", str(s.get(key, "")))
                txt.pack(side="left", fill="x", expand=True)
                self.set_vars[key] = txt
                # Text widgets don't expose trace; bind focusout + key release.
                txt.bind("<FocusOut>", lambda _e: self._auto_save_full_settings())
                txt.bind("<KeyRelease>", lambda _e: self._auto_save_full_settings())
            else:
                v = tk.StringVar(value=str(s.get(key, "")))
                self.set_vars[key] = v
                ttk.Entry(row, textvariable=v, width=60).pack(side="left", fill="x", expand=True)
                v.trace_add("write", lambda *_: self._auto_save_full_settings())

        def add_int(key, label):
            row = tk.Frame(inner, bg="#ffffff"); row.pack(fill="x", pady=2)
            tk.Label(row, text=label, bg="#ffffff", width=22, anchor="w").pack(side="left")
            v = tk.IntVar(value=int(s.get(key, 0) or 0))
            self.set_vars[key] = v
            ttk.Spinbox(row, textvariable=v, from_=0, to=99999, width=10).pack(side="left")
            v.trace_add("write", lambda *_: self._auto_save_full_settings())

        def add_check(key, label):
            v = tk.BooleanVar(value=bool(s.get(key, False)))
            self.set_checks[key] = v
            tk.Checkbutton(inner, text=label, variable=v, bg="#ffffff",
                           activebackground="#ffffff",
                           command=self._auto_save_full_settings).pack(
                anchor="w", padx=4, pady=1,
            )

        add_section("标题与副本")
        add_check("title_prefix_enabled", "启用标题前缀")
        add_str("title_prefix", "标题前缀：")
        add_check("title_random_suffix_enabled", "启用标题随机后缀")
        add_int("random_suffix_length", "随机后缀长度：")
        add_int("output_copies", "每产品输出份数：")
        add_check("split_output_files", "将每份输出拆成单独 xlsx")
        add_section("品牌与类目")
        add_check("brand_enabled", "启用品牌覆盖")
        add_str("brand_value", "品牌：")
        add_check("category_enabled", "启用类目覆盖")
        add_str("category_value", "TikTok 类目：")
        add_section("价格 / 库存 / COD")
        add_check("price_enabled", "启用价格覆盖")
        add_int("price_value", "价格 (PHP)：")
        add_check("quantity_enabled", "启用库存覆盖")
        add_int("quantity_value", "库存数量：")
        add_check("cod_enabled", "启用 COD 覆盖")
        add_str("cod_value", "COD：")
        add_section("尺码")
        add_check("fill_sizes_enabled", "用标准尺码覆盖 property_value_2")
        add_str("standard_sizes", "标准尺码：")
        add_section("描述与尺码图")
        add_check("description_enabled", "启用描述覆盖")
        add_str("description_value", "产品描述：", multi=True)
        add_check("size_chart_enabled", "启用尺码图 URL 覆盖")
        add_str("size_chart_value", "尺码图 URL / ID：")
        add_section("包裹尺寸 (cm / g)")
        add_check("parcel_enabled", "启用包裹尺寸覆盖")
        add_int("parcel_weight_value", "重量（克）：")
        add_int("parcel_length_value", "长 (cm)：")
        add_int("parcel_width_value", "宽 (cm)：")
        add_int("parcel_height_value", "高 (cm)：")
        add_section("源表列名映射（按需调整）")
        self.colmap_vars = {}
        mapping = self.cfg.get("source_column_mapping") or default_config()["source_column_mapping"]
        for field_name, default_col in mapping.items():
            row = tk.Frame(inner, bg="#ffffff"); row.pack(fill="x", pady=1)
            tk.Label(row, text=field_name, bg="#ffffff", width=22, anchor="w").pack(side="left")
            v = tk.StringVar(value=str(default_col or ""))
            self.colmap_vars[field_name] = v
            ttk.Entry(row, textvariable=v, width=40).pack(side="left", fill="x", expand=True)
            v.trace_add("write", lambda *_: self._auto_save_full_settings())

        ttk.Button(inner, text="保存设置", command=self._save_full_settings).pack(anchor="e", pady=12)

    # ----- Log + status chrome -------------------------------------------

    def _build_log_panel(self, parent: tk.Misc):
        container = tk.Frame(parent, bg="#ffffff")
        container.pack(side="bottom", fill="x", padx=0, pady=0)

        title = tk.Label(container, text="日志", bg="#ffffff",
                         fg="#222", font=("", 10, "bold"))
        title.pack(anchor="w", padx=20, pady=(8, 0))

        log_frame = tk.Frame(container, bg="#ffffff")
        log_frame.pack(fill="x", padx=20, pady=(2, 4))
        self._log_widget = scrolledtext.ScrolledText(
            log_frame, height=10, wrap="word",
            font=("Consolas", 9), bg="#fafbfc", fg="#222",
            bd=1, relief="solid",
        )
        self._log_widget.pack(fill="x")
        self._log_widget.configure(state="disabled")

        bar = tk.Frame(container, bg="#ffffff")
        bar.pack(fill="x", padx=20, pady=(0, 8))
        self.var_status = tk.StringVar(value="就绪")
        tk.Label(bar, textvariable=self.var_status, bg="#ffffff",
                 fg="#444", anchor="w").pack(side="left")
        self.progress = ttk.Progressbar(bar, mode="determinate", maximum=1000, length=300)
        self.progress.pack(side="right")
        tk.Label(bar, text=f"v{cfg_mod.__version__}", bg="#ffffff",
                 fg="#888", font=("", 9)).pack(side="right", padx=(0, 12))

    # ----- Source picker + drag/drop -------------------------------------

    def _pick_source(self):
        p = filedialog.askopenfilename(
            title="选择源表 xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if p:
            self._on_source_dropped(p)

    def _on_source_dropped(self, path: str):
        self.var_source.set(path)
        self.cfg["product_xlsx_last"] = path
        save_config(self.cfg)
        # Auto-detect format and counts.
        try:
            fmt = detect_format(path)
            products = read_source(path, column_mapping=self.cfg.get("source_column_mapping"))
            variants = sum(len(p.variants) for p in products)
            self._source_meta = {
                "format": fmt,
                "products": len(products),
                "skus": variants,
            }
            label = {
                "easyboss": "EasyBoss 导出#SKU",
                "tiktok":   "TikTok 批量上传模板",
            }.get(fmt, "未识别")
            self._var_source_summary.set(
                f"源文件:  {label}  ·  产品 {len(products)} 个 · SKU {variants} 行"
            )
            self.append_log(f"[scan] {path} → 格式 {label}，{len(products)} 产品 / {variants} SKU")
        except Exception as e:  # noqa: BLE001
            self._var_source_summary.set(f"源文件: 无法识别 - {e}")
            self.append_log(f"[scan] {path} → 识别失败: {e}")

    def _pick_output(self):
        initial = self.var_output.get() or str(get_app_dir())
        p = filedialog.askdirectory(title="选择输出目录", initialdir=initial)
        if p:
            self.var_output.set(p)
            self._auto_save_paths()

    def _auto_save_paths(self):
        self.cfg["product_xlsx_last"] = self.var_source.get()
        self.cfg["product_xlsx_output_dir"] = self.var_output.get()
        self.cfg["product_pool_dir"] = self.var_pool.get() if hasattr(self, "var_pool") else ""
        save_config(self.cfg)

    def _auto_save_settings(self):
        # Collect from the compact panel and merge.
        new = self._convert_panel.collect()
        s = self.cfg.setdefault("product_xlsx_settings", {})
        for k, v in new.items():
            s[k] = v
        # Also save source path so it sticks across runs.
        self.cfg["product_xlsx_last"] = self.var_source.get()
        self.cfg["product_xlsx_output_dir"] = self.var_output.get()
        save_config(self.cfg)

    # ----- Conversion driver ---------------------------------------------

    def _do_convert(self):
        src = self.var_source.get().strip()
        if not src or not Path(src).exists():
            messagebox.showwarning("提示", "请先在「文件」里选一个有效的 xlsx。")
            return
        out = self._resolve_output_dir()
        out.mkdir(parents=True, exist_ok=True)
        tpl = template_path()
        if not tpl.exists():
            messagebox.showerror("模板缺失", f"找不到模板文件：\n{tpl}\n请确认 assets/ 目录完整。")
            return
        # Pull in everything from the panel — including the just-edited fields.
        self._auto_save_settings()
        settings = get_settings(self.cfg)
        mapping = self.cfg.get("source_column_mapping") or {}
        # v3.2.4: 转化页可勾选跳过图片下载（默认勾上 = 下载）
        download_imgs = bool(
            self._convert_panel.is_download_images_enabled()
            if hasattr(self, "_convert_panel") else True
        )
        self.append_log(f"[convert] 源={src}\n[convert] 输出={out}\n[convert] 下载图片={'是' if download_imgs else '否'}")
        self.worker.submit(
            convert_source,
            source_xlsx=Path(src),
            output_dir=out,
            settings=settings,
            template_src=tpl,
            column_mapping=mapping,
            download_imgs=download_imgs,
            on_done=self._on_convert_done,
        )

    def _on_convert_done(self, result):
        if result is None:
            return
        self.cfg["product_xlsx_last_output_dir"] = str(result.output_path.parent)
        save_config(self.cfg)
        if len(result.output_paths) > 1:
            lines = "\n".join(f"  • {p.name}" for p in result.output_paths)
            self.append_log(
                f"[convert] 成功 · 产品 {result.product_count} / 行 {result.row_count}\n"
                f"[convert] 已拆分为 {len(result.output_paths)} 个文件:\n{lines}"
            )
            messagebox.showinfo(
                "转化完成",
                f"已生成 {result.product_count} 个产品，"
                f"共 {len(result.output_paths)} 个文件 / {result.row_count} 行。\n\n"
                f"输出文件：\n" + "\n".join(str(p) for p in result.output_paths),
            )
            self._open_in_explorer(result.output_paths[0].parent)
        else:
            self.append_log(
                f"[convert] 成功 · 产品 {result.product_count} / 行 {result.row_count} → "
                f"{result.output_path.name}"
            )
            messagebox.showinfo(
                "转化完成",
                f"已生成 {result.row_count} 行 / {result.product_count} 个产品。\n\n"
                f"输出文件：\n{result.output_path}",
            )
            self._open_in_explorer(result.output_path)

    # ----- Pool driver ----------------------------------------------------

    def _do_add_pool(self):
        src = self.var_source.get().strip()
        if not src or not Path(src).exists():
            messagebox.showwarning("提示", "请先选择有效的源表 xlsx。")
            return
        pool_dir = Path(self.var_output.get().strip() or (get_app_dir() / "产品池"))
        name = self.var_pool_name.get().strip() or "默认池"
        mapping = self.cfg.get("source_column_mapping") or {}
        self.worker.submit(
            add_to_pool,
            pool_dir=pool_dir,
            pool_name=name,
            source_xlsx=Path(src),
            column_mapping=mapping,
            on_done=self._on_add_done,
        )

    def _on_add_done(self, info):
        if info is None:
            return
        self._refresh_pools()
        self.append_log(f"[pool] 已加入「{info.name}」，共 {info.count} 行。")
        messagebox.showinfo("入池完成", f"已加入「{info.name}」，共 {info.count} 行。")

    def _do_extract_convert(self):
        sel = self.tree_pools.selection() if hasattr(self, "tree_pools") else ()
        if not sel:
            messagebox.showwarning("提示", "请先在列表里选择一个产品池。")
            return
        item = self.tree_pools.item(sel[0])
        filename = item["values"][2]
        pool_dir = Path(self.var_output.get().strip() or (get_app_dir() / "产品池"))
        pool_path = pool_dir / str(filename)
        if not pool_path.exists():
            messagebox.showerror("错误", f"找不到池文件：\n{pool_path}")
            return
        out = self._resolve_output_dir()
        tpl = template_path()
        settings = get_settings(self.cfg)
        mapping = self.cfg.get("source_column_mapping") or {}

        def task(progress, log, **_):
            progress(0.05, f"读取产品池 {pool_path.name}…")
            products = load_pool(pool_path)
            if not products:
                raise ValueError("产品池为空或无法识别。")
            from .converter import generate_random_suffix
            from .tiktok_writer import build_rows_for_product, write_tiktok_xlsx
            copies = max(1, int(settings.get("output_copies", 1) or 1))
            split_files = bool(settings.get("split_output_files", False))
            # Output path: one file (default) or N tagged files.
            import time as _t
            ts = _t.strftime("%Y%m%d_%H%M%S")
            stem = Path(str(filename)).stem
            output_paths: list[Path] = []
            total_rows = 0

            if split_files:
                from .converter import generate_copy_suffixes
                copies_suffixes = (
                    generate_copy_suffixes(copies, int(settings.get("random_suffix_length", 10)))
                    if settings.get("title_random_suffix_enabled") else [""] * copies
                )
                for c_idx in range(copies):
                    file_rows: list = []
                    for p in products:
                        file_rows.extend(
                            build_rows_for_product(
                                p, settings,
                                copy_suffixes=[copies_suffixes[c_idx]],
                                apply_suffix_to_first=True,
                            )
                        )
                    tag = f"copy{c_idx + 1:02d}of{copies:02d}" if copies > 1 else "copy01of01"
                    out_path = out / f"{stem}_TKPH_{ts}_{tag}.xlsx"
                    write_tiktok_xlsx(out_path, file_rows, tpl)
                    output_paths.append(out_path)
                    total_rows += len(file_rows)
            else:
                copies_suffixes = (
                    generate_copy_suffixes(copies, int(settings.get("random_suffix_length", 10)))
                    if settings.get("title_random_suffix_enabled") else [""] * copies
                )
                rows = []
                for i, p in enumerate(products, 1):
                    rows.extend(build_rows_for_product(p, settings, copy_suffixes=copies_suffixes))
                    progress(0.2 + 0.6 * (i / len(products)), f"已处理 {i}/{len(products)}…")
                out_path = out / f"{stem}_TKPH_{ts}.xlsx"
                write_tiktok_xlsx(out_path, rows, tpl)
                output_paths.append(out_path)
                total_rows = len(rows)

            progress(1.0, f"完成：{len(products)} 产品 / {total_rows} 行 → {output_paths[-1].name}")
            from .converter import ConvertResult
            return ConvertResult(
                output_path=output_paths[0],
                output_paths=output_paths,
                product_count=len(products),
                row_count=total_rows,
            )

        self.worker.submit(task, on_done=self._on_convert_done)

    # ----- Settings save --------------------------------------------------

    def _save_full_settings(self, silent: bool = False):
        """Save every Settings-tab field back into ``self.cfg`` and to disk.

        ``silent=True`` is the auto-save path used by per-widget trace_add
        callbacks; suppress the success dialog and any error noise so we
        don't ruin the user's flow with popups for every keystroke.
        """
        s = self.cfg.setdefault("product_xlsx_settings", {})
        for k, v in self.set_vars.items():
            if isinstance(v, tk.Text):
                s[k] = v.get("1.0", "end-1c")
            else:
                try:
                    s[k] = v.get()
                except tk.TclError:
                    s[k] = ""
        for k, v in self.set_checks.items():
            s[k] = bool(v.get())
        if hasattr(self, "colmap_vars"):
            self.cfg["source_column_mapping"] = {
                k: v.get() for k, v in self.colmap_vars.items()
            }
        try:
            save_config(self.cfg)
        except OSError as e:
            if not silent:
                messagebox.showerror("保存失败", f"写入 config.json 失败：\n{e}")
            raise
        if not silent:
            messagebox.showinfo("已保存", "设置已写入 config.json。")

    def _auto_save_full_settings(self):
        """Silent version: used by trace_add callbacks on Settings-tab fields."""
        try:
            self._save_full_settings(silent=True)
        except OSError:
            pass  # already warned by the explicit save path

    # ----- Pool list ------------------------------------------------------

    def _refresh_pools(self):
        if not hasattr(self, "tree_pools"):
            return
        pool_dir = Path(self.var_output.get().strip() or (get_app_dir() / "产品池"))
        for row in self.tree_pools.get_children():
            self.tree_pools.delete(row)
        try:
            pools = list_pools(pool_dir)
        except Exception as e:  # noqa: BLE001
            self.var_status.set(f"读取产品池失败：{e}")
            return
        for p in pools:
            self.tree_pools.insert("", "end", values=(p.name, p.count, p.filename))
        self.var_status.set(
            f"产品池目录：{pool_dir}（{len(pools)} 个池）"
            if pools else f"产品池目录为空：{pool_dir}"
        )

    # ----- Progress / busy / log ----------------------------------------

    def on_progress(self, p: float, text: str):
        self.progress["value"] = max(0, min(1000, p * 1000))
        self.var_status.set(text)

    def append_log(self, text: str):
        self._log_widget.configure(state="normal")
        self._log_widget.insert("end", text + "\n")
        self._log_widget.see("end")
        self._log_widget.configure(state="disabled")

    def set_busy(self, busy: bool):
        try:
            self.configure(cursor="watch" if busy else "")
        except tk.TclError:
            pass

    def _resolve_output_dir(self) -> Path:
        out = self.var_output.get().strip()
        if out:
            return Path(out)
        src = self.var_source.get().strip()
        base = Path(src).parent if src else get_app_dir()
        return base / "ready"

    def _open_in_explorer(self, path: Path):
        try:
            if os.sys.platform == "darwin":
                os.system(f'open "{path}"')
            elif os.name == "nt":
                os.startfile(str(path))  # type: ignore[attr-defined]
            else:
                os.system(f'xdg-open "{path}"')
        except Exception:  # noqa: BLE001
            pass


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
