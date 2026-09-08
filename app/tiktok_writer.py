"""Write product rows into the TikTok Shop Philippines batch upload template.

The template (assets/batch-product-source.xlsx) has 12 sheets. We only write
data into the 'Template' sheet starting at row 2, leaving the rest of the
workbook (Instruction, Image, Example, HiddenStyle, HiddenAttr, etc.)
untouched so that TikTok's validator still accepts the file.

ROW STRATEGY (v3 — TikTok-compatible):
    The EasyBoss source table has one row per (color × size) SKU. We expand
    each SKU into its own TikTok listing row (each row = one variant of one
    product). This matches TikTok's official Example sheet where every
    (color, size) combination is a separate listing.

    Common fields (description, size_chart, parcel, brand, category, images,
    delivery, cod, pre_order_time) are shared across all rows of the same
    product.

    `output_copies` (防查重) duplicates each variant N times with a unique
    random suffix in the title and seller_sku, so duplicate listings can be
    detected and removed later.

    When `fill_sizes_enabled` is on, the writer falls back to the standard
    sizes list (S,M,L,XL,2XL,3XL) ONLY when the source product has no
    variation-2 values (single-SKU products). Otherwise each row keeps its
    own size value, and the field gets exactly one size per row.
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import openpyxl

from .config import get_assets_dir
from .source_reader import Product, Variant


TEMPLATE_FILENAME = "batch-product-source.xlsx"

# The 39 Template-sheet column headers, in order
TIKTOK_COLUMNS: list[str] = [
    "category",
    "brand",
    "product_name",
    "product_description",
    "main_image",
    "image_2", "image_3", "image_4", "image_5",
    "image_6", "image_7", "image_8", "image_9",
    "property_name_1",
    "property_value_1",
    "property_1_image",
    "property_name_2",
    "property_value_2",
    "parcel_weight",
    "parcel_length",
    "parcel_width",
    "parcel_height",
    "delivery",
    "price",
    "pre_order_time",
    "quantity",
    "seller_sku",
    "size_chart",
    "cod",
    "product_property/100157",
    "product_property/100198",
    "product_property/100393",
    "product_property/100395",
    "product_property/100397",
    "product_property/100398",
    "product_property/100399",
    "product_property/100400",
    "product_property/100401",
    "product_property/100403",
]


def template_path() -> Path:
    """Locate the template inside the app's assets/ directory."""
    return get_assets_dir() / TEMPLATE_FILENAME


def _kg_to_grams(v: Any) -> Any:
    """EasyBoss source stores weight in KG; TikTok wants grams."""
    if v is None or v == "":
        return None
    try:
        f = float(v)
        return int(round(f * 1000))
    except (TypeError, ValueError):
        return v


def _clean_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _to_number(v: Any) -> Any:
    """Coerce a value to int/float so Excel stores it as a number.

    EasyBoss sometimes stores numeric columns as text (e.g. '356' instead of
    356). TikTok's validator and the import pipeline require numeric types
    for price / quantity / parcel dimensions / weight.
    """
    if v is None or v == "":
        return None
    if isinstance(v, bool):  # bool is a subclass of int — treat False/True as 0/1
        return int(v)
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if not s:
        return None
    # Strip common currency / unit markers
    for ch in [",", " ", "₱", "$", "￥", "¥", "PHP", "php", "kg", "KG", "g", "G", "cm", "CM"]:
        s = s.replace(ch, "")
    try:
        f = float(s)
        if f.is_integer():
            return int(f)
        return f
    except ValueError:
        return v  # leave as-is


# Variation name translations: Chinese → English (TikTok backend expects English).
VAR_NAME_TRANSLATIONS: dict[str, str] = {
    # Chinese
    "颜色": "Color",
    "顏色": "Color",
    "尺码": "Size",
    "尺碼": "Size",
    "尺寸": "Size",
    "规格": "Specification",
    # English-already is a pass-through (we just normalize capitalization)
    "color": "Color",
    "colour": "Color",
    "size": "Size",
}


def _translate_var_name(name: str) -> str:
    """Translate a variation name (Chinese/English) to TikTok-expected English."""
    if not name:
        return ""
    key = _clean_str(name)
    if key in VAR_NAME_TRANSLATIONS:
        return VAR_NAME_TRANSLATIONS[key]
    # Common case: pascal/title case English → leave as-is
    return key


# A single output row: a dict {col_name: value}
OutputRow = dict[str, Any]


def _resolve_common_fields(
    product: Product,
    settings: dict[str, Any],
) -> dict[str, Any]:
    """Resolve fields that are shared across all rows of a product."""
    brand = (
        settings["brand_value"]
        if settings.get("brand_enabled") and settings.get("brand_value")
        else (product.brand or "")
    )
    if not brand:
        brand = "No brand"

    price = (
        settings["price_value"]
        if settings.get("price_enabled")
        else None  # leave to per-variant price if not configured
    )
    quantity = (
        settings["quantity_value"]
        if settings.get("quantity_enabled")
        else None
    )
    cod = settings["cod_value"] if settings.get("cod_enabled") else ""
    category = (
        settings["category_value"]
        if settings.get("category_enabled")
        else product.category
    )
    description = (
        settings["description_value"]
        if settings.get("description_enabled")
        else product.description
    )
    size_chart = (
        settings["size_chart_value"]
        if settings.get("size_chart_enabled")
        else product.size_chart
    )

    if settings.get("parcel_enabled"):
        weight = settings.get("parcel_weight_value")
        length = settings.get("parcel_length_value")
        width = settings.get("parcel_width_value")
        height = settings.get("parcel_height_value")
    else:
        weight = _kg_to_grams(product.parcel_weight_kg)
        length = product.parcel_length
        width = product.parcel_width
        height = product.parcel_height

    return {
        "brand": _clean_str(brand),
        "category": _clean_str(category),
        "description": _clean_str(description),
        "size_chart": _clean_str(size_chart),
        "weight": weight,
        "length": length,
        "width": width,
        "height": height,
        "cod": cod,
        "quantity": quantity,
        "price_override": price,
        "delivery": _clean_str(settings.get("delivery_value", "")),
        "pre_order_time": settings.get("pre_order_time_value", ""),
    }


def _build_row_for_variant(
    product: Product,
    variant: Variant,
    settings: dict[str, Any],
    common: dict[str, Any],
    copy_idx: int = 0,
    copy_suffix: str = "",
) -> OutputRow:
    """Build a single TikTok row for one (variant, copy) combination."""
    # Title
    title_prefix = settings["title_prefix"] if settings.get("title_prefix_enabled") else ""
    suffix = copy_suffix if copy_idx > 0 else ""
    title = f"{title_prefix}{product.product_name}{suffix}".strip()

    # Variant-level values
    var1_name = _translate_var_name(product.var1_name or "颜色")
    var1_value = _clean_str(variant.var1)
    var2_name = _translate_var_name(product.var2_name or "尺码")

    if variant.var2:
        var2_value = _clean_str(variant.var2)
    elif settings.get("fill_sizes_enabled"):
        # Fallback: source has no size; use the standard sizes list
        std = settings.get("standard_sizes") or "S,M,L,XL,2XL,3XL"
        var2_value = std
    else:
        var2_value = ""

    # Images: prefer variant's SKU image for the main; fall back to product images
    var1_image = _clean_str(variant.sku_image)
    images = list(product.images[:9])
    while len(images) < 9:
        images.append("")
    images = [_clean_str(i) for i in images]

    # Price: prefer per-variant price, then override, then None
    price_value = _to_number(variant.price) if variant.price not in (None, "") else _to_number(common["price_override"])

    # Quantity: prefer per-variant stock, then override
    quantity_value = (
        _to_number(variant.stock) if variant.stock not in (None, "") else _to_number(common["quantity"])
    )

    # Seller SKU: use full platform_sku (unique per variant)
    base_sku = _clean_str(variant.platform_sku) or product.master_sku()
    seller_sku = base_sku + copy_suffix if copy_suffix and copy_idx > 0 else base_sku

    row: OutputRow = {col: "" for col in TIKTOK_COLUMNS}
    row["category"] = common["category"]
    row["brand"] = common["brand"]
    row["product_name"] = title
    row["product_description"] = common["description"]
    # main_image: prefer variant SKU image, else first product image
    row["main_image"] = var1_image if var1_image else images[0] if images else ""
    for i, img in enumerate(images[:9]):
        col = "main_image" if i == 0 else f"image_{i + 1}"
        if i == 0 and var1_image:
            # already set above
            continue
        row[col] = img
    row["property_name_1"] = var1_name
    row["property_value_1"] = var1_value
    row["property_1_image"] = var1_image
    row["property_name_2"] = var2_name
    row["property_value_2"] = var2_value
    row["parcel_weight"] = _to_number(common["weight"])
    row["parcel_length"] = _to_number(common["length"])
    row["parcel_width"] = _to_number(common["width"])
    row["parcel_height"] = _to_number(common["height"])
    row["delivery"] = common["delivery"]
    row["price"] = price_value
    row["pre_order_time"] = common["pre_order_time"]
    row["quantity"] = quantity_value
    row["seller_sku"] = seller_sku
    row["size_chart"] = common["size_chart"]
    row["cod"] = common["cod"]
    return row


def build_rows_for_product(
    product: Product,
    settings: dict[str, Any],
    copy_suffixes: list[str] | None = None,
) -> list[OutputRow]:
    """Build all TikTok rows for one Product.

    Each variant becomes one row; if `output_copies` > 1, each variant is
    duplicated `output_copies` times (防查重). Random suffixes are passed in
    via `copy_suffixes` (length = output_copies). The first copy uses an
    empty suffix.
    """
    copies = max(1, int(settings.get("output_copies", 1) or 1))
    if copy_suffixes is None:
        copy_suffixes = [""] * copies
    elif len(copy_suffixes) < copies:
        copy_suffixes = (copy_suffixes + [""] * copies)[: copies]

    common = _resolve_common_fields(product, settings)
    rows: list[OutputRow] = []

    if not product.variants:
        # Edge case: product with no variants → one row, no variation values
        for c_idx in range(copies):
            rows.append(_build_row_for_variant(
                product, Variant(), settings, common,
                copy_idx=c_idx, copy_suffix=copy_suffixes[c_idx],
            ))
        return rows

    for v in product.variants:
        for c_idx in range(copies):
            rows.append(_build_row_for_variant(
                product, v, settings, common,
                copy_idx=c_idx, copy_suffix=copy_suffixes[c_idx],
            ))
    return rows


def write_tiktok_xlsx(
    output_path: Path,
    rows: list[OutputRow],
    template_src: Path,
) -> Path:
    """Copy the template to output_path, then write `rows` into the Template sheet."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_src, output_path)
    wb = openpyxl.load_workbook(output_path)
    if "Template" not in wb.sheetnames:
        wb.close()
        raise ValueError(f"模板文件缺少 'Template' sheet：{template_src}")
    ws = wb["Template"]

    # Build header→col index from the first row of the Template sheet
    header_row = [(_clean_str(c.value) or "") for c in ws[1]]
    col_idx: dict[str, int] = {}
    for i, h in enumerate(header_row, 1):
        if h in TIKTOK_COLUMNS:
            col_idx[h] = i

    # Find first data row (the template's data area usually starts at row 2;
    # skip any pre-existing instruction rows by locating the header row)
    start_row = 2
    # Clear any pre-existing data rows in the Template sheet
    max_existing = ws.max_row
    if max_existing >= start_row:
        ws.delete_rows(start_row, max_existing - start_row + 1)

    # Write new rows
    for r_off, row_dict in enumerate(rows):
        excel_row = start_row + r_off
        for col_name, value in row_dict.items():
            ci = col_idx.get(col_name)
            if ci is None:
                continue
            ws.cell(row=excel_row, column=ci, value=value)

    wb.save(output_path)
    wb.close()
    return output_path