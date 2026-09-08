"""Tkinter GUI for the TK Philippines table converter.

Layout:
  Top  : three path rows (source xlsx, output dir, product-pool dir)
  Mid  : ttk.Notebook with tabs [转化 / 入池 / 提取 / 设置]
  Bot  : status label + progress bar

All long-running work runs on a background thread; UI updates are
posted back via a thread-safe queue and drained on the Tk event loop
so the interface never freezes.
"""
from __future__ import annotations

import os
import queue
import threading
import traceback
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import config as cfg_mod
from .config import (
    APP_NAME,
    default_config,
    get_app_dir,
    get_settings,
    load_config,
    save_config,
    update_settings,
)
from .converter import convert_source
from .product_pool import (
    PoolInfo,
    add_to_pool,
    list_pools,
    load_pool,
)
from .source_reader import detect_format
from .tiktok_writer import template_path


# ---------------------------------------------------------------------------
# Worker bridge
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
                result = fn(*args, progress=progress, **kwargs)
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
                elif kind == "done":
                    _, result, on_done = item
                    self.app.set_busy(False)
                    if on_done:
                        on_done(result)
                elif kind == "error":
                    _, exc, tb, on_done = item
                    self.app.set_busy(False)
                    self.app.on_progress(0, f"错误: {exc}")
                    messagebox.showerror("运行出错", f"{exc}\n\n{tb}")
                    if on_done:
                        on_done(None)
        except queue.Empty:
            pass
        if self._t and self._t.is_alive():
            self.app.after(80, self._drain)


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME + " v" + cfg_mod.__version__)
        self.geometry("960x720")
        self.minsize(880, 640)

        self.cfg = load_config()
        self.worker = Worker(self)

        # Tk variables bound to the three top paths
        self.var_source = tk.StringVar(value=self.cfg.get("product_xlsx_last", ""))
        self.var_output = tk.StringVar(value=self.cfg.get("product_xlsx_output_dir", ""))
        self.var_pool = tk.StringVar(value=self.cfg.get("product_pool_dir", ""))

        self._build_ui()
        self._bind_traces()
        self._refresh_pools()

    # ----- UI construction -------------------------------------------------

    def _build_ui(self):
        try:
            ttk.Style(self).theme_use("clam")
        except tk.TclError:
            pass

        # ---- Top: three path rows
        top = ttk.LabelFrame(self, text="路径设置")
        top.pack(side="top", fill="x", padx=10, pady=(10, 6))

        for i, (label, var, kind) in enumerate([
            ("源表 xlsx：", self.var_source, "open_xlsx"),
            ("输出目录：", self.var_output, "dir"),
            ("产品池目录：", self.var_pool, "dir"),
        ]):
            ttk.Label(top, text=label).grid(row=i, column=0, sticky="w", padx=8, pady=4)
            ent = ttk.Entry(top, textvariable=var)
            ent.grid(row=i, column=1, sticky="ew", padx=4, pady=4)
            ttk.Button(top, text="选择…", width=10,
                       command=lambda v=var, k=kind: self._pick(v, k)
                       ).grid(row=i, column=2, padx=8)
        top.columnconfigure(1, weight=1)

        # ---- Middle: notebook
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=6)
        self._build_tab_convert()
        self._build_tab_add()
        self._build_tab_extract()
        self._build_tab_settings()

        # ---- Bottom: status + progress
        bot = ttk.Frame(self)
        bot.pack(side="bottom", fill="x", padx=10, pady=(0, 10))
        self.var_status = tk.StringVar(value="就绪")
        ttk.Label(bot, textvariable=self.var_status, anchor="w").pack(side="top", fill="x")
        self.progress = ttk.Progressbar(bot, mode="determinate", maximum=1000)
        self.progress.pack(side="top", fill="x", pady=(4, 0))

    # ----- Convert tab ----------------------------------------------------

    def _build_tab_convert(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text="转化")
        info = (
            "1. 选择 EasyBoss 导出 #SKU 表格（或 TikTok 批量上传模板本身）。\n"
            "2. 在「设置」里调整标题前缀、品牌、价格、尺码、包裹等。\n"
            "3. 点「开始转化」，结果会保存到「输出目录」下的 ready/ 子目录。\n"
        )
        ttk.Label(f, text=info, justify="left").pack(anchor="w", padx=12, pady=(12, 6))
        btn_row = ttk.Frame(f); btn_row.pack(fill="x", padx=12, pady=6)
        ttk.Button(btn_row, text="开始转化", command=self._do_convert).pack(side="left")
        ttk.Button(btn_row, text="识别源表格式", command=self._do_detect).pack(side="left", padx=8)
        self.var_format = tk.StringVar(value="")
        ttk.Label(btn_row, textvariable=self.var_format, foreground="#0a6").pack(side="left", padx=8)

    def _do_detect(self):
        p = self.var_source.get().strip()
        if not p or not Path(p).exists():
            messagebox.showwarning("提示", "请先在「源表」里选一个有效的 xlsx。")
            return
        try:
            fmt = detect_format(p)
            label = {"easyboss": "EasyBoss 导出#SKU", "tiktok": "TikTok 批量上传模板"}.get(fmt, "未识别")
            self.var_format.set(f"已识别：{label}")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("识别失败", str(e))

    def _do_convert(self):
        src = self.var_source.get().strip()
        if not src or not Path(src).exists():
            messagebox.showwarning("提示", "请先选择有效的源表 xlsx。")
            return
        out = self._resolve_output_dir()
        out.mkdir(parents=True, exist_ok=True)
        tpl = template_path()
        if not tpl.exists():
            messagebox.showerror("模板缺失", f"找不到模板文件：\n{tpl}\n请确认 assets/ 目录完整。")
            return
        settings = get_settings(self.cfg)
        mapping = self.cfg.get("source_column_mapping") or {}
        self.worker.submit(
            convert_source,
            source_xlsx=Path(src),
            output_dir=out,
            settings=settings,
            template_src=tpl,
            column_mapping=mapping,
            on_done=self._on_convert_done,
        )

    def _on_convert_done(self, result):
        if result is None:
            return
        self.cfg["product_xlsx_last_output_dir"] = str(result.output_path.parent)
        save_config(self.cfg)
        messagebox.showinfo(
            "转化完成",
            f"已生成 {result.row_count} 行 / {result.product_count} 个产品。\n\n"
            f"输出文件：\n{result.output_path}",
        )
        self._open_in_explorer(result.output_path)

    # ----- Add-to-pool tab ------------------------------------------------

    def _build_tab_add(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text="入池")
        ttk.Label(f, text="把当前源表里的产品加入指定的产品池。\n"
                          "同一个池可以多次入池，工具会自动追加。",
                  justify="left").pack(anchor="w", padx=12, pady=(12, 6))
        row = ttk.Frame(f); row.pack(fill="x", padx=12, pady=6)
        ttk.Label(row, text="池名称：").pack(side="left")
        self.var_pool_name = tk.StringVar(value="默认池")
        ttk.Entry(row, textvariable=self.var_pool_name, width=24).pack(side="left", padx=4)
        ttk.Button(row, text="加入产品池", command=self._do_add_pool).pack(side="left", padx=8)

        ttk.Separator(f).pack(fill="x", padx=12, pady=8)
        ttk.Label(f, text="当前产品池中的池：").pack(anchor="w", padx=12)
        cols = ("name", "count", "updated")
        self.tree_pools = ttk.Treeview(f, columns=cols, show="headings", height=8)
        self.tree_pools.heading("name", text="池名称")
        self.tree_pools.heading("count", text="行数")
        self.tree_pools.heading("updated", text="更新时间")
        self.tree_pools.column("name", width=220, anchor="w")
        self.tree_pools.column("count", width=80, anchor="e")
        self.tree_pools.column("updated", width=180, anchor="w")
        self.tree_pools.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        ttk.Button(f, text="刷新", command=self._refresh_pools).pack(anchor="e", padx=12, pady=(0, 8))

    def _do_add_pool(self):
        src = self.var_source.get().strip()
        if not src or not Path(src).exists():
            messagebox.showwarning("提示", "请先选择有效的源表 xlsx。")
            return
        pool_dir = Path(self.var_pool.get().strip()) if self.var_pool.get().strip() else (get_app_dir() / "产品池")
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
        messagebox.showinfo("入池完成", f"已加入「{info.name}」，共 {info.count} 行。")

    # ----- Extract tab ----------------------------------------------------

    def _build_tab_extract(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text="提取")
        ttk.Label(f, text="从产品池中提取产品，可以直接再次转化为 TikTok 模板。",
                  justify="left").pack(anchor="w", padx=12, pady=(12, 6))
        top = ttk.Frame(f); top.pack(fill="both", expand=True, padx=12, pady=6)
        cols = ("name", "count", "filename")
        self.tree_extract = ttk.Treeview(top, columns=cols, show="headings", selectmode="browse")
        self.tree_extract.heading("name", text="池名称")
        self.tree_extract.heading("count", text="行数")
        self.tree_extract.heading("filename", text="文件")
        self.tree_extract.column("name", width=200, anchor="w")
        self.tree_extract.column("count", width=80, anchor="e")
        self.tree_extract.column("filename", width=260, anchor="w")
        self.tree_extract.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(top, orient="vertical", command=self.tree_extract.yview)
        self.tree_extract.configure(yscrollcommand=sb.set); sb.pack(side="left", fill="y")

        btns = ttk.Frame(f); btns.pack(fill="x", padx=12, pady=6)
        ttk.Button(btns, text="刷新", command=self._refresh_pools).pack(side="left")
        ttk.Button(btns, text="提取并转化 →", command=self._do_extract_convert).pack(side="right")

    def _do_extract_convert(self):
        sel = self.tree_extract.selection()
        if not sel:
            messagebox.showwarning("提示", "请先在列表里选择一个产品池。")
            return
        item = self.tree_extract.item(sel[0])
        filename = item["values"][2]
        pool_dir = Path(self.var_pool.get().strip()) if self.var_pool.get().strip() else (get_app_dir() / "产品池")
        pool_path = pool_dir / str(filename)
        if not pool_path.exists():
            messagebox.showerror("错误", f"找不到池文件：\n{pool_path}")
            return
        # Read products then convert
        out = self._resolve_output_dir()
        tpl = template_path()
        settings = get_settings(self.cfg)
        mapping = self.cfg.get("source_column_mapping") or {}

        def task(progress, **_):
            progress(0.05, f"读取产品池 {pool_path.name}…")
            products = load_pool(pool_path)
            if not products:
                raise ValueError("产品池为空或无法识别。")
            progress(0.2, f"已识别 {len(products)} 个产品，正在转化…")
            from .converter import generate_random_suffix
            from .tiktok_writer import build_rows_for_product, write_tiktok_xlsx
            rows = []
            for i, p in enumerate(products, 1):
                rs = generate_random_suffix(int(settings.get("random_suffix_length", 3))) \
                    if settings.get("title_random_suffix_enabled") else ""
                rows.extend(build_rows_for_product(p, settings, random_suffix=rs))
                progress(0.2 + 0.6 * (i / len(products)), f"已处理 {i}/{len(products)}…")
            import time as _t
            ts = _t.strftime("%Y%m%d_%H%M%S")
            stem = Path(str(filename)).stem
            out_path = out / f"{stem}_TKPH_{ts}.xlsx"
            write_tiktok_xlsx(out_path, rows, tpl)
            progress(1.0, f"完成：{len(products)} 产品 / {len(rows)} 行 → {out_path.name}")
            from .converter import ConvertResult
            return ConvertResult(output_path=out_path, product_count=len(products), row_count=len(rows))

        self.worker.submit(task, on_done=self._on_convert_done)

    # ----- Settings tab ---------------------------------------------------

    def _build_tab_settings(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text="设置")
        # Scrollable canvas
        canvas = tk.Canvas(f, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(f, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = ttk.Frame(canvas)
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(inner_id, width=e.width))

        s = get_settings(self.cfg)
        self.set_vars: dict[str, tk.Variable] = {}
        self.set_checks: dict[str, tk.BooleanVar] = {}

        def add_section(title: str):
            ttk.Label(inner, text=title, font=("", 11, "bold")).pack(anchor="w", padx=12, pady=(12, 4))

        def add_str(key: str, label: str, width: int = 60, multi: bool = False):
            row = ttk.Frame(inner); row.pack(fill="x", padx=12, pady=2)
            ttk.Label(row, text=label, width=22, anchor="w").pack(side="left")
            v = tk.StringVar(value=str(s.get(key, "")))
            self.set_vars[key] = v
            if multi:
                txt = tk.Text(row, height=6, width=width, wrap="word")
                txt.insert("1.0", v.get())
                txt.pack(side="left", fill="x", expand=True)
                self.set_vars[key] = txt  # store widget for multi-line
            else:
                ent = ttk.Entry(row, textvariable=v, width=width)
                ent.pack(side="left", fill="x", expand=True)
            return v

        def add_int(key: str, label: str):
            row = ttk.Frame(inner); row.pack(fill="x", padx=12, pady=2)
            ttk.Label(row, text=label, width=22, anchor="w").pack(side="left")
            v = tk.IntVar(value=int(s.get(key, 0) or 0))
            self.set_vars[key] = v
            ttk.Spinbox(row, textvariable=v, from_=0, to=99999, width=10).pack(side="left")
            return v

        def add_check(key: str, label: str, on_frame: ttk.Frame | None = None):
            parent = on_frame or inner
            v = tk.BooleanVar(value=bool(s.get(key, False)))
            self.set_checks[key] = v
            ttk.Checkbutton(parent, text=label, variable=v).pack(anchor="w", padx=12, pady=1)
            return v

        # 标题
        add_section("标题与副本")
        add_check("title_prefix_enabled", "启用标题前缀")
        add_str("title_prefix", "标题前缀：")
        add_check("title_random_suffix_enabled", "启用标题随机后缀（用于多 listing）")
        add_int("random_suffix_length", "随机后缀长度：")
        add_int("output_copies", "每产品输出份数：")

        # 品牌 / 类目
        add_section("品牌与类目")
        add_check("brand_enabled", "启用品牌覆盖")
        add_str("brand_value", "品牌：")
        add_check("category_enabled", "启用类目覆盖")
        add_str("category_value", "TikTok 类目（如 Men's Tops/T-shirts）：")

        # 价格 / 库存 / COD
        add_section("价格 / 库存 / COD")
        add_check("price_enabled", "启用价格覆盖")
        add_int("price_value", "价格 (PHP)：")
        add_check("quantity_enabled", "启用库存覆盖")
        add_int("quantity_value", "库存数量：")
        add_check("cod_enabled", "启用 COD 覆盖")
        add_str("cod_value", "COD（Y/N）：")

        # 尺码
        add_section("尺码")
        add_check("fill_sizes_enabled", "用标准尺码覆盖 property_value_2")
        add_str("standard_sizes", "标准尺码（逗号分隔）：")

        # 描述 / 尺码图
        add_section("描述与尺码图")
        add_check("description_enabled", "启用描述覆盖")
        add_str("description_value", "产品描述：", width=60, multi=True)
        add_check("size_chart_enabled", "启用尺码图 URL 覆盖")
        add_str("size_chart_value", "尺码图 URL：")

        # 包裹
        add_section("包裹尺寸 (cm / g)")
        add_check("parcel_enabled", "启用包裹尺寸覆盖")
        add_int("parcel_weight_value", "重量（克）：")
        add_int("parcel_length_value", "长 (cm)：")
        add_int("parcel_width_value", "宽 (cm)：")
        add_int("parcel_height_value", "高 (cm)：")

        # TikTok 特定
        add_section("TikTok 特定")
        add_int("pre_order_time_value", "pre_order_time（天）：")
        add_str("delivery_value", "delivery（发货方式）：")

        # 源列映射
        add_section("源表列名映射 (EasyBoss 默认，可按需调整)")
        self._build_column_mapping_widgets(inner, s)

        # 保存按钮
        ttk.Button(inner, text="保存设置", command=self._save_settings).pack(anchor="e", padx=12, pady=12)

    def _build_column_mapping_widgets(self, parent, s):
        self.colmap_vars: dict[str, tk.StringVar] = {}
        mapping = self.cfg.get("source_column_mapping") or default_config()["source_column_mapping"]
        for field_name, default_col in mapping.items():
            row = ttk.Frame(parent); row.pack(fill="x", padx=12, pady=1)
            ttk.Label(row, text=field_name, width=22, anchor="w").pack(side="left")
            v = tk.StringVar(value=str(default_col or ""))
            self.colmap_vars[field_name] = v
            ttk.Entry(row, textvariable=v, width=40).pack(side="left", fill="x", expand=True)

    def _save_settings(self):
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
        # column mapping
        cm = {k: v.get() for k, v in self.colmap_vars.items()}
        self.cfg["source_column_mapping"] = cm
        save_config(self.cfg)
        messagebox.showinfo("已保存", "设置已写入 config.json。")

    # ----- Pool listing ----------------------------------------------------

    def _refresh_pools(self):
        pool_dir = Path(self.var_pool.get().strip()) if self.var_pool.get().strip() else (get_app_dir() / "产品池")
        for tree in (getattr(self, "tree_pools", None), getattr(self, "tree_extract", None)):
            if tree is None:
                continue
            for row in tree.get_children():
                tree.delete(row)
        try:
            pools = list_pools(pool_dir)
        except Exception as e:  # noqa: BLE001
            self.var_status.set(f"读取产品池失败：{e}")
            return
        for p in pools:
            if self.tree_pools:
                self.tree_pools.insert("", "end", values=(p.name, p.count, p.filename))
            if self.tree_extract:
                self.tree_extract.insert("", "end", values=(p.name, p.count, p.filename))
        if not pools:
            self.var_status.set(f"产品池目录为空：{pool_dir}")
        else:
            self.var_status.set(f"已加载 {len(pools)} 个产品池。")

    # ----- Progress / state ------------------------------------------------

    def on_progress(self, p: float, text: str):
        self.progress["value"] = max(0, min(1000, p * 1000))
        self.var_status.set(text)

    def set_busy(self, busy: bool):
        try:
            self.configure(cursor="watch" if busy else "")
        except tk.TclError:
            pass

    def _pick(self, var: tk.StringVar, kind: str):
        if kind == "open_xlsx":
            p = filedialog.askopenfilename(
                title="选择源表 xlsx",
                filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
            )
        else:
            initial = var.get() or str(get_app_dir())
            p = filedialog.askdirectory(title="选择目录", initialdir=initial)
        if p:
            var.set(p)

    def _resolve_output_dir(self) -> Path:
        out = self.var_output.get().strip()
        if out:
            p = Path(out)
        else:
            # default: source xlsx's directory / ready
            src = self.var_source.get().strip()
            base = Path(src).parent if src else get_app_dir()
            p = base / "ready"
        return p

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

    def _bind_traces(self):
        for v in (self.var_source, self.var_output, self.var_pool):
            v.trace_add("write", lambda *_: self._on_paths_change())

    def _on_paths_change(self):
        self.cfg["product_xlsx_last"] = self.var_source.get()
        self.cfg["product_xlsx_output_dir"] = self.var_output.get()
        self.cfg["product_pool_dir"] = self.var_pool.get()
        # Don't save on every keystroke; save on convert / quit
        if hasattr(self, "tree_pools"):
            self._refresh_pools()


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
