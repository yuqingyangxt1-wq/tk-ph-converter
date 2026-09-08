"""Read source tables: EasyBoss 导出#SKU or TikTok batch upload template.

Auto-detects the format by header keywords, then maps columns to the internal
Product model. EasyBoss source has one row per SKU/variant — we group by
product name and merge variants into a single Product.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import openpyxl


# ---------------------------------------------------------------------------
# Internal data model
# ---------------------------------------------------------------------------


@dataclass
class Variant:
    """A single SKU row from the source (one combination of variations)."""
    platform_sku: str = ""
    var1: str = ""
    var2: str = ""
    var3: str = ""
    price: Any = None
    stock: Any = None
    sku_image: str = ""


@dataclass
class Product:
    """A master product, aggregated from 1+ source rows."""
    product_name: str = ""
    brand: str = ""
    category: str = ""
    description: str = ""
    size_chart: str = ""
    images: list[str] = field(default_factory=list)        # up to 9
    parcel_weight_kg: Any = None
    parcel_length: Any = None
    parcel_width: Any = None
    parcel_height: Any = None
    var1_name: str = ""
    var2_name: str = ""
    var3_name: str = ""
    variants: list[Variant] = field(default_factory=list)

    # Filled by reader
    source_format: str = "unknown"  # "easyboss" or "tiktok"
    source_row_count: int = 0

    # ------------------------------------------------------------------
    # Derived helpers
    # ------------------------------------------------------------------
    def unique_var_values(self, which: int) -> list[str]:
        """Return unique, ordered, non-empty values for variation 1/2/3."""
        seen: list[str] = []
        for v in self.variants:
            val = (getattr(v, f"var{which}", "") or "").strip()
            if val and val not in seen:
                seen.append(val)
        return seen

    def var_values_str(self, which: int) -> str:
        return ",".join(self.unique_var_values(which))

    def master_sku(self) -> str:
        """Best-effort master SKU: strip the last 2 '-suffix' segments from
        the first variant's platform_sku.  E.g.
        '1635504750_45715669615-Black-S' → '1635504750_45715669615'.

        The master portion often contains underscores (it is a numeric id
        pair), so we can't use '_' as a separator.  We DO require that the
        last two segments look like variation values — short tokens without
        whitespace and without a digit-only pattern that would make the
        master ambiguous.
        """
        if not self.variants:
            return ""
        first = self.variants[0].platform_sku or ""
        if "-" not in first:
            return first
        parts = first.rsplit("-", 2)
        if len(parts) != 3 or not parts[1] or not parts[2]:
            return first
        # Heuristic: the master is the long numeric/underscore prefix; the
        # last two tokens are short (variation values like Black, S, M).
        # If the "master" part is shorter than either suffix, give up.
        if len(parts[0]) < max(len(parts[1]), len(parts[2])):
            return first
        return parts[0]


# ---------------------------------------------------------------------------
# Header auto-detection
# ---------------------------------------------------------------------------


# A row whose headers contain ALL these tokens → TikTok batch upload template
TIKTOK_HEADER_TOKENS = (
    "product_name",
    "main_image",
    "property_name_1",
    "parcel_weight",
    "pre_order_time",
    "seller_sku",
)


# EasyBoss 导出#SKU header synonyms (Chinese + English-ish).
# Each maps a logical field → set of possible header strings.
EASYBOSS_HEADER_SYNONYMS: dict[str, set[str]] = {
    "product_name":    {"产品名", "商品名称", "产品名称", "商品标题", "标题", "Product Name", "product_name", "Title", "title"},
    "brand":           {"品牌", "Brand", "brand"},
    "category":        {"产品类目", "商品类目", "类目", "类目路径", "Category", "category"},
    "platform_sku":    {"平台SKU", "平台 SKU", "平台sku", "SKU", "sku", "商品SKU", "Seller SKU", "seller_sku"},
    "var1_name":       {"规格1名称", "规格 1 名称", "变体名1", "Variation 1 Name", "Primary variation name"},
    "var1_value":      {"规格1选项", "规格 1 选项", "变体值1", "Variation 1 Value", "Primary variation value"},
    "var2_name":       {"规格2名称", "规格 2 名称", "变体名2", "Variation 2 Name", "Secondary variation name"},
    "var2_value":      {"规格2选项", "规格 2 选项", "变体值2", "Variation 2 Value", "Secondary variation value"},
    "var3_name":       {"规格3名称", "规格 3 名称", "变体名3"},
    "var3_value":      {"规格3选项", "规格 3 选项", "变体值3"},
    "price":           {"税前价格", "价格", "售价", "Price", "price", "Retail Price"},
    "stock":           {"库存", "数量", "Quantity", "quantity", "Stock", "stock"},
    "sku_image":       {"SKU图片", "SKU 图片", "变体图片", "Primary variation image"},
    "image_1":         {"产品图片1", "主图", "商品主图", "Main image", "main_image"},
    "image_2":         {"产品图片2", "Image 2", "image_2"},
    "image_3":         {"产品图片3", "Image 3", "image_3"},
    "image_4":         {"产品图片4", "Image 4", "image_4"},
    "image_5":         {"产品图片5", "Image 5", "image_5"},
    "image_6":         {"产品图片6", "Image 6", "image_6"},
    "image_7":         {"产品图片7", "Image 7", "image_7"},
    "image_8":         {"产品图片8", "Product Image 8", "image_8"},
    "image_9":         {"产品图片9", "Product Image 9", "image_9"},
    "parcel_weight_kg":{"包裹重量（KG）", "包裹重量(KG)", "包裹重量", "重量", "Package weight(kg)", "parcel_weight"},
    "parcel_length":   {"包裹长度（CM）", "包裹长度(CM)", "包裹长度", "Package length(cm)", "parcel_length"},
    "parcel_width":    {"包裹宽度（CM）", "包裹宽度(CM)", "包裹宽度", "Package width(cm)", "parcel_width"},
    "parcel_height":   {"包裹高度（CM）", "包裹高度(CM)", "包裹高度", "Package height(cm)", "parcel_height"},
    "description":     {"产品描述", "商品描述", "描述", "Description", "product_description"},
    "size_chart":      {"尺码图", "Size Chart", "size_chart"},
}


def _build_header_index(headers: list[str]) -> dict[str, int]:
    """Map normalized header string → column index (0-based)."""
    return {(h or "").strip(): i for i, h in enumerate(headers)}


def _detect_format(headers: list[str]) -> str:
    norm = [h.strip() for h in headers]
    hits = sum(1 for t in TIKTOK_HEADER_TOKENS if t in norm)
    if hits >= 4:
        return "tiktok"
    # Heuristic: if any EasyBoss-specific Chinese header is present
    if any("产品名" in h or "平台SKU" in h or "规格" in h or "税前价格" in h for h in norm):
        return "easyboss"
    return "unknown"


def _resolve_mapping(
    headers: list[str],
    override: dict[str, str] | None,
) -> dict[str, int]:
    """Resolve logical field → column index.

    1. Use the override mapping (configured column name → field).
    2. If override is missing/blank, fall back to synonym matching.
    """
    index = _build_header_index(headers)
    out: dict[str, int] = {}
    for field_name, synonyms in EASYBOSS_HEADER_SYNONYMS.items():
        col: int | None = None
        if override and override.get(field_name):
            wanted = override[field_name].strip()
            if wanted in index:
                col = index[wanted]
        if col is None:
            for syn in synonyms:
                if syn in index:
                    col = index[syn]
                    break
        if col is not None:
            out[field_name] = col
    return out


def _row_get(row: tuple, col_idx: int | None) -> Any:
    if col_idx is None or col_idx >= len(row):
        return None
    return row[col_idx]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def read_source(
    xlsx_path: str | Path,
    column_mapping: dict[str, str] | None = None,
) -> list[Product]:
    """Read a source xlsx and return a list of Products (grouped by product name)."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    headers = list(rows[0])
    fmt = _detect_format(headers)
    if fmt == "unknown":
        raise ValueError(
            "无法识别表格格式：表头既不像 EasyBoss 导出#SKU，也不像 TikTok 批量上传模板。"
        )

    products: dict[str, Product] = {}  # grouped by product_name (and brand/category)

    if fmt == "easyboss":
        mapping = _resolve_mapping(headers, column_mapping)
        for row in rows[1:]:
            name = (_row_get(row, mapping.get("product_name")) or "").strip() or ""
            if not name:
                continue
            key = name  # group key
            p = products.get(key)
            if p is None:
                p = Product(
                    product_name=name,
                    brand=str(_row_get(row, mapping.get("brand")) or "").strip(),
                    category=str(_row_get(row, mapping.get("category")) or "").strip(),
                    description=str(_row_get(row, mapping.get("description")) or ""),
                    size_chart=str(_row_get(row, mapping.get("size_chart")) or "").strip(),
                    parcel_weight_kg=_row_get(row, mapping.get("parcel_weight_kg")),
                    parcel_length=_row_get(row, mapping.get("parcel_length")),
                    parcel_width=_row_get(row, mapping.get("parcel_width")),
                    parcel_height=_row_get(row, mapping.get("parcel_height")),
                    var1_name=str(_row_get(row, mapping.get("var1_name")) or "").strip(),
                    var2_name=str(_row_get(row, mapping.get("var2_name")) or "").strip(),
                    var3_name=str(_row_get(row, mapping.get("var3_name")) or "").strip(),
                    source_format="easyboss",
                )
                # images: product image 1..9
                for i in range(1, 10):
                    v = _row_get(row, mapping.get(f"image_{i}"))
                    if v:
                        p.images.append(str(v).strip())
                products[key] = p

            # variant
            v = Variant(
                platform_sku=str(_row_get(row, mapping.get("platform_sku")) or "").strip(),
                var1=str(_row_get(row, mapping.get("var1_value")) or "").strip(),
                var2=str(_row_get(row, mapping.get("var2_value")) or "").strip(),
                var3=str(_row_get(row, mapping.get("var3_value")) or "").strip(),
                price=_row_get(row, mapping.get("price")),
                stock=_row_get(row, mapping.get("stock")),
                sku_image=str(_row_get(row, mapping.get("sku_image")) or "").strip(),
            )
            p.variants.append(v)
            p.source_row_count += 1

    else:  # tiktok
        # Build an index by the exact TikTok header names
        idx = _build_header_index(headers)
        def tget(row, key):
            ci = idx.get(key)
            return _row_get(row, ci)

        for row in rows[1:]:
            name = (tget(row, "product_name") or "").strip()
            if not name:
                continue
            if name in products:
                # Already have this product; just bump counter
                products[name].source_row_count += 1
                continue
            p = Product(
                product_name=name,
                brand=str(tget(row, "brand") or "").strip(),
                category=str(tget(row, "category") or "").strip(),
                description=str(tget(row, "product_description") or ""),
                size_chart=str(tget(row, "size_chart") or "").strip(),
                parcel_weight_kg=tget(row, "parcel_weight"),  # grams already
                parcel_length=tget(row, "parcel_length"),
                parcel_width=tget(row, "parcel_width"),
                parcel_height=tget(row, "parcel_height"),
                var1_name=str(tget(row, "property_name_1") or "").strip(),
                var2_name=str(tget(row, "property_name_2") or "").strip(),
                var3_name="",
                source_format="tiktok",
            )
            for key in ("main_image", "image_2", "image_3", "image_4",
                        "image_5", "image_6", "image_7", "image_8", "image_9"):
                v = tget(row, key)
                if v:
                    p.images.append(str(v).strip())
            seller_sku = tget(row, "seller_sku")
            v = Variant(platform_sku=str(seller_sku or "").strip())
            p.variants.append(v)
            p.source_row_count += 1
            products[name] = p

    wb.close()
    return list(products.values())


def detect_format(xlsx_path: str | Path) -> str:
    """Cheap format probe (only reads the first row)."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    wb.close()
    if not first:
        return "unknown"
    return _detect_format(list(first))
