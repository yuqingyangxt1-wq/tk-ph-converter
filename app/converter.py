"""Core conversion orchestration: source → Products → TikTok rows → xlsx."""
from __future__ import annotations

import random
import string
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from .source_reader import Product, read_source
from .tiktok_writer import build_rows_for_product, write_tiktok_xlsx


ProgressFn = Callable[[float, str], None]   # (0.0..1.0, status_text)


def generate_random_suffix(length: int = 3) -> str:
    """Generate a short random token appended to titles to differentiate copies."""
    length = max(1, min(int(length or 3), 8))
    chars = string.ascii_uppercase + string.digits
    # Avoid leading character that looks like O/I/0/1 confusion
    safe_first = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    first = random.choice(safe_first)
    rest = "".join(random.choices(chars, k=length - 1))
    return first + rest


@dataclass
class ConvertResult:
    output_path: Path
    product_count: int
    row_count: int


def convert_source(
    source_xlsx: Path,
    output_dir: Path,
    settings: dict[str, Any],
    template_src: Path,
    progress: ProgressFn | None = None,
    column_mapping: dict[str, str] | None = None,
) -> ConvertResult:
    """Read source, apply settings, write TikTok batch upload xlsx.

    Output filename: <source-stem>_TKPH_<timestamp>.xlsx in `output_dir`.
    """
    if progress:
        progress(0.02, "正在读取源表格…")
    products = read_source(source_xlsx, column_mapping=column_mapping)
    if not products:
        raise ValueError("源表格中没有可识别的产品数据。")

    if progress:
        progress(0.25, f"已识别 {len(products)} 个产品，正在应用设置…")
    rows: list[dict[str, Any]] = []
    total = len(products)
    for i, p in enumerate(products, 1):
        # Unique random suffix per *product copy-set* to keep title pairs grouped
        from .tiktok_writer import build_rows_for_product as _br  # noqa: F401
        if settings.get("title_random_suffix_enabled"):
            rs = generate_random_suffix(int(settings.get("random_suffix_length", 3)))
        else:
            rs = ""
        rows.extend(build_rows_for_product(p, settings, random_suffix=rs))
        if progress and total:
            progress(0.25 + 0.5 * (i / total), f"已处理 {i}/{total} 个产品…")
        # Tiny yield so the UI thread can repaint
        if i % 10 == 0:
            time.sleep(0)

    if progress:
        progress(0.8, "正在写入 TikTok 模板…")
    output_dir.mkdir(parents=True, exist_ok=True)
    ts = time.strftime("%Y%m%d_%H%M%S")
    out_path = output_dir / f"{source_xlsx.stem}_TKPH_{ts}.xlsx"
    write_tiktok_xlsx(out_path, rows, template_src)

    if progress:
        progress(1.0, f"完成。共 {len(products)} 个产品 / {len(rows)} 行 → {out_path.name}")
    return ConvertResult(output_path=out_path, product_count=len(products), row_count=len(rows))
