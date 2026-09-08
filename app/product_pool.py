"""Product pool: 入池 (add) and 提取 (extract) from a directory of xlsx files.

Layout (matches the original tool's 产品池 folder):
  <pool_dir>/
    _pools_meta.json     ← index of pools
    <pool_name>.xlsx     ← one xlsx per pool, each row = one master product
                           (stored using the same EasyBoss header format so
                           it round-trips back through the reader)
    把产品池表格放这里.txt
"""
from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import openpyxl

from .source_reader import Product, read_source


ProgressFn = Callable[[float, str], None]


META_FILENAME = "_pools_meta.json"
PLACEHOLDER_TXT = "把产品池表格放这里.txt"


# ---------------------------------------------------------------------------
# Pool index (meta)
# ---------------------------------------------------------------------------


@dataclass
class PoolInfo:
    name: str           # display name (also the xlsx stem)
    filename: str       # the actual xlsx filename
    count: int          # product count (cached in meta)
    updated: str        # last update timestamp
    source_format: str  # "easyboss" (round-trip) or "tiktok"

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "filename": self.filename,
            "count": self.count,
            "updated": self.updated,
            "source_format": self.source_format,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "PoolInfo":
        return cls(
            name=d.get("name", ""),
            filename=d.get("filename", ""),
            count=int(d.get("count", 0) or 0),
            updated=d.get("updated", ""),
            source_format=d.get("source_format", "easyboss"),
        )


def _load_meta(pool_dir: Path) -> dict[str, Any]:
    p = pool_dir / META_FILENAME
    if p.exists():
        try:
            with p.open("r", encoding="utf-8") as f:
                return json.load(f)
        except (OSError, json.JSONDecodeError):
            pass
    return {"version": 1, "pools": []}


def _save_meta(pool_dir: Path, meta: dict[str, Any]) -> None:
    pool_dir.mkdir(parents=True, exist_ok=True)
    with (pool_dir / META_FILENAME).open("w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


def _ensure_pool_dir(pool_dir: Path) -> Path:
    pool_dir.mkdir(parents=True, exist_ok=True)
    placeholder = pool_dir / PLACEHOLDER_TXT
    if not placeholder.exists():
        placeholder.write_text("把产品池表格放这里\n", encoding="utf-8")
    return pool_dir


def _safe_filename(name: str) -> str:
    """Sanitize a pool name into a valid xlsx filename."""
    name = (name or "").strip() or "默认池"
    # Replace Windows-illegal characters
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    if not name.lower().endswith(".xlsx"):
        name = name + ".xlsx"
    return name


# Canonical EasyBoss header order — we store pool rows in this order so
# they can be re-read by source_reader and re-converted.
POOL_HEADERS: list[str] = [
    "产品名", "站点", "店铺ID", "店铺名称", "品牌", "产品类目", "平台SKU",
    "规格1名称", "规格1选项", "规格2名称", "规格2选项", "规格3名称", "规格3选项",
    "税前价格", "库存", "SKU图片",
    "产品图片1", "产品图片2", "产品图片3", "产品图片4", "产品图片5",
    "产品图片6", "产品图片7", "产品图片8", "产品图片9",
    "包裹重量（KG）", "包裹长度（CM）", "包裹宽度（CM）", "包裹高度（CM）",
    "仓库", "产品描述", "尺码图",
]


def _product_to_pool_rows(p: Product, site: str = "PH") -> list[list[Any]]:
    """Expand a Product back into one or more EasyBoss-format rows.

    One product → 1+ rows (one per variant) so the round-trip preserves
    the original row granularity.
    """
    rows: list[list[Any]] = []
    images = list(p.images[:9])
    while len(images) < 9:
        images.append("")
    # If the product has no variants, still emit at least one row
    variants = p.variants or [None]
    for v in variants:
        row: list[Any] = [""] * len(POOL_HEADERS)
        row[0] = p.product_name
        row[1] = site
        row[4] = p.brand
        row[5] = p.category
        if v is not None:
            row[6] = v.platform_sku
            row[7] = p.var1_name
            row[8] = v.var1
            row[9] = p.var2_name
            row[10] = v.var2
            row[11] = p.var3_name
            row[12] = v.var3
            row[13] = v.price
            row[14] = v.stock
            row[15] = v.sku_image
        for i, img in enumerate(images):
            row[16 + i] = img
        row[25] = p.parcel_weight_kg
        row[26] = p.parcel_length
        row[27] = p.parcel_width
        row[28] = p.parcel_height
        row[30] = p.description
        row[31] = p.size_chart
        rows.append(row)
    return rows


def _load_existing_pool_rows(path: Path) -> list[list[Any]]:
    if not path.exists():
        return []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    # Drop header row if present
    if rows and rows[0] and rows[0][0] == POOL_HEADERS[0]:
        rows = rows[1:]
    return rows


def add_to_pool(
    pool_dir: Path,
    pool_name: str,
    source_xlsx: Path,
    progress: ProgressFn | None = None,
    column_mapping: dict[str, str] | None = None,
    site: str = "PH",
) -> PoolInfo:
    """入池: read source xlsx, group into Products, append to the named pool xlsx."""
    _ensure_pool_dir(pool_dir)
    products = read_source(source_xlsx, column_mapping=column_mapping)
    if not products:
        raise ValueError("源表格中没有可识别的产品数据。")
    if progress:
        progress(0.2, f"已识别 {len(products)} 个产品，准备入池…")

    fname = _safe_filename(pool_name)
    pool_path = pool_dir / fname

    existing = _load_existing_pool_rows(pool_path)
    new_rows: list[list[Any]] = []
    for i, p in enumerate(products, 1):
        new_rows.extend(_product_to_pool_rows(p, site=site))
        if progress:
            progress(0.2 + 0.6 * (i / len(products)), f"已展开 {i}/{len(products)} 个产品…")

    all_rows = existing + new_rows
    if progress:
        progress(0.85, f"正在写入 {pool_path.name}…")
    _write_pool_xlsx(pool_path, all_rows)

    info = PoolInfo(
        name=pool_name.strip() or Path(fname).stem,
        filename=fname,
        count=len(all_rows),
        updated=time.strftime("%Y-%m-%d %H:%M:%S"),
        source_format="easyboss",
    )
    _update_meta_entry(pool_dir, info)
    if progress:
        progress(1.0, f"已入池 {len(products)} 个产品到「{pool_name}」，累计 {len(all_rows)} 行。")
    return info


def _write_pool_xlsx(path: Path, rows: list[list[Any]]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(POOL_HEADERS)
    for r in rows:
        ws.append(r)
    wb.save(path)
    wb.close()


def _update_meta_entry(pool_dir: Path, info: PoolInfo) -> None:
    meta = _load_meta(pool_dir)
    pools = meta.get("pools", [])
    pools = [p for p in pools if p.get("filename") != info.filename]
    pools.append(info.to_dict())
    # Sort by updated desc
    pools.sort(key=lambda p: p.get("updated", ""), reverse=True)
    meta["pools"] = pools
    _save_meta(pool_dir, meta)


def list_pools(pool_dir: Path) -> list[PoolInfo]:
    _ensure_pool_dir(pool_dir)
    meta = _load_meta(pool_dir)
    out: list[PoolInfo] = []
    for p in meta.get("pools", []):
        try:
            out.append(PoolInfo.from_dict(p))
        except Exception:
            continue
    # Also pick up any orphan xlsx files not in the meta
    known = {p.filename for p in out}
    for x in sorted(pool_dir.glob("*.xlsx")):
        if x.name in known:
            continue
        try:
            wb = openpyxl.load_workbook(x, data_only=True, read_only=True)
            ws = wb.worksheets[0]
            n = max(0, (ws.max_row or 0) - 1)
            wb.close()
        except Exception:
            n = 0
        out.append(PoolInfo(name=x.stem, filename=x.name, count=n,
                            updated="", source_format="easyboss"))
    return out


def load_pool(pool_path: Path) -> list[Product]:
    """提取: read a pool xlsx back into Products."""
    return read_source(pool_path)
