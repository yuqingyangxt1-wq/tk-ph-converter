"""Write product rows into the TikTok Shop Philippines batch upload template.

The template (assets/batch-product-source.xlsx) has 12 sheets. We only write
data into the 'Template' sheet starting at row 2, leaving the rest of the
workbook (Instruction, Image, Example, HiddenStyle, HiddenAttr, etc.)
untouched so that TikTok's validator still accepts the file.
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter

from .config import get_assets_dir
from .source_reader import Product


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


# A single output row: a dict {col_name: value}
OutputRow = dict[str, Any]


def build_rows_for_product(
    product: Product,
    settings: dict[str, Any],
    random_suffix: str = "",
) -> list[OutputRow]:
    """Apply settings to a Product and return 1+ TikTok template rows.

    Each call returns `output_copies` rows (if enabled) with a different
    random suffix appended to the title.
    """
    copies = max(1, int(settings.get("output_copies", 1) or 1))

    # Resolve common values once
    title_prefix = settings["title_prefix"] if settings["title_prefix_enabled"] else ""
    brand = settings["brand_value"] if settings["brand_enabled"] else (product.brand or "")
    price = settings["price_value"] if settings["price_enabled"] else None
    quantity = settings["quantity_value"] if settings["quantity_enabled"] else None
    cod = settings["cod_value"] if settings["cod_enabled"] else ""
    category = settings["category_value"] if settings["category_enabled"] else product.category
    description = settings["description_value"] if settings["description_enabled"] else product.description
    size_chart = settings["size_chart_value"] if settings["size_chart_enabled"] else product.size_chart

    if settings["parcel_enabled"]:
        weight = settings["parcel_weight_value"]
        length = settings["parcel_length_value"]
        width = settings["parcel_width_value"]
        height = settings["parcel_height_value"]
    else:
        weight = _kg_to_grams(product.parcel_weight_kg)
        length = product.parcel_length
        width = product.parcel_width
        height = product.parcel_height

    # Variations: fill_sizes overrides variation 2 with standard sizes
    if settings["fill_sizes_enabled"]:
        std = (settings.get("standard_sizes") or "S,M,L,XL,2XL,3XL")
        var2_name = product.var2_name or "Size"
        var2_value = std
    else:
        var2_name = product.var2_name
        var2_value = product.var_values_str(2)

    var1_name = product.var1_name
    var1_value = product.var_values_str(1)
    var1_image = product.variants[0].sku_image if product.variants else ""

    base_title = product.product_name
    images = list(product.images[:9])  # main + image_2..9
    while len(images) < 9:
        images.append("")

    rows: list[OutputRow] = []
    for copy_idx in range(copies):
        # Random suffix
        suffix = ""
        if settings.get("title_random_suffix_enabled"):
            if copy_idx == 0 and not random_suffix:
                # Caller didn't pre-generate; build one
                from .converter import generate_random_suffix  # avoid circular at import
                suffix = generate_random_suffix(int(settings.get("random_suffix_length", 3)))
            else:
                suffix = random_suffix
        # Title assembly: prefix + base + (if fill_sizes, append size hint? — keep base; sizes go to variation) + suffix
        # We do NOT inject sizes into the title text — they go into property_value_2.
        # But a short size tag can be appended if desired; the original config uses
        # "【S-3XL】" already in the prefix, so we keep titles text-only here.
        title = f"{title_prefix}{base_title}{suffix}".strip()

        row: OutputRow = {col: "" for col in TIKTOK_COLUMNS}
        row["category"] = _clean_str(category)
        row["brand"] = _clean_str(brand)
        row["product_name"] = title
        row["product_description"] = _clean_str(description)
        for i, img in enumerate(images[:9]):
            col = "main_image" if i == 0 else f"image_{i+1}"
            row[col] = img
        row["property_name_1"] = _clean_str(var1_name)
        row["property_value_1"] = var1_value
        row["property_1_image"] = var1_image
        row["property_name_2"] = _clean_str(var2_name)
        row["property_value_2"] = var2_value
        row["parcel_weight"] = weight
        row["parcel_length"] = length
        row["parcel_width"] = width
        row["parcel_height"] = height
        row["delivery"] = _clean_str(settings.get("delivery_value", ""))
        row["price"] = price
        row["pre_order_time"] = settings.get("pre_order_time_value", "")
        row["quantity"] = quantity
        row["seller_sku"] = product.master_sku() + (suffix if suffix and copy_idx > 0 else "")
        row["size_chart"] = size_chart
        row["cod"] = cod
        # product_property/* columns left blank; user can fill post-export
        rows.append(row)
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
