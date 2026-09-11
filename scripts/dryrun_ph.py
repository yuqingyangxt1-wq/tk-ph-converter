"""Dry-run PH tool v3.2.3 against a real EasyBoss PH xlsx source table.

Same logic as scripts/dryrun_th.py in the TH repo — just imports PH's app/,
uses PH's default settings, and runs convert_source() against an arbitrary
source xlsx passed as the first argument (or the bundled mock fallback).
"""
from __future__ import annotations
import sys, shutil
from pathlib import Path

# Defaults: assume this script lives in tk-ph-converter-repo/scripts/
HERE = Path(__file__).resolve().parent
REPO = HERE.parent
sys.path.insert(0, str(REPO))

import openpyxl
from app.config import default_config
from app.converter import convert_source


def main(src_arg: str | None) -> None:
    if src_arg:
        mock_src = Path(src_arg).resolve()
        if not mock_src.exists():
            raise SystemExit(f"源表不存在: {mock_src}")
        print(f"using real source: {mock_src}  ({mock_src.stat().st_size:,} bytes)")
    else:
        raise SystemExit("usage: dryrun.py <source.xlsx>")

    out_dir = REPO / "dryrun_out"
    if out_dir.exists():
        shutil.rmtree(out_dir)

    logs: list[str] = []
    def log(s): logs.append(s)
    def progress(p, msg): logs.append(f"  progress {p:.2f}  {msg}")

    cfg = default_config()
    ps = cfg["product_xlsx_settings"]

    print()
    print("--- DEFAULT_SETTINGS['product_xlsx_settings'] (PH) ---")
    for k in ("title_prefix", "price_value", "category_value", "brand_value",
              "output_copies", "size_chart_value", "fill_sizes_enabled",
              "split_output_files", "parcel_weight_value"):
        print(f"  {k} = {ps.get(k)!r}")

    template = REPO / "assets" / "batch-product-source.xlsx"
    if not template.exists():
        raise SystemExit(f"PH template missing: {template}")
    print(f"\nusing template: {template}  ({template.stat().st_size:,} bytes)")

    print("\n>>> running convert_source() ...")
    result = convert_source(
        source_xlsx=mock_src,
        output_dir=out_dir,
        settings=ps,
        template_src=template,
        progress=progress,
        log=log,
        download_imgs=False,  # skip; the URLs are Shopee PH CDN and will 403 outside PH anyway
    )

    print(f"\n--- CONVERT RESULT ---")
    print(f"  output_paths: {result.output_paths}")
    print(f"  products: {result.product_count}  rows: {result.row_count}")

    # Inspect the output xlsx
    out_xlsx = result.output_paths[0]
    print(f"\n--- INSPECT {out_xlsx.name} ---")
    wb2 = openpyxl.load_workbook(out_xlsx)
    ws2 = wb2["Template"]
    headers_out = [str(c.value or "") for c in ws2[1]]
    print(f"  output column count: {len(headers_out)}")
    print(f"  pre_order_time present? {'pre_order_time' in headers_out}")

    # Show first 3 rows + last 1 row
    print(f"\n--- FIRST 3 ROWS ---")
    for r in range(2, min(5, ws2.max_row + 1)):
        cells = []
        for c in range(1, ws2.max_column + 1):
            v = ws2.cell(r, c).value
            if v is not None and v != "":
                label = ws2.cell(1, c).value
                s = str(v)
                if len(s) > 40:
                    s = s[:37] + "..."
                cells.append(f"[{label}]={s}")
        print(f"  R{r}: {' | '.join(cells)}")

    print(f"\n--- LAST ROW (R{ws2.max_row}) ---")
    r = ws2.max_row
    cells = []
    for c in range(1, ws2.max_column + 1):
        v = ws2.cell(r, c).value
        if v is not None and v != "":
            label = ws2.cell(1, c).value
            s = str(v)
            if len(s) > 40:
                s = s[:37] + "..."
            cells.append(f"[{label}]={s}")
    print(f"  R{r}: {' | '.join(cells)}")

    # Field spot-check on first data row
    def cell(label):
        idx = headers_out.index(label) + 1
        return ws2.cell(2, idx).value

    print(f"\n--- KEY FIELD CHECK (row 2) ---")
    checks = [
        ("brand", "No brand"),
        ("category", "Men's Tops/T-shirts"),  # PH default
        ("price", 356),  # PH default
        ("parcel_weight", 200),  # 0.2 KG → 200g
        ("property_name_1", "Color"),
        ("property_name_2", "Size"),
    ]
    for label, expected in checks:
        v = cell(label)
        ok = (v == expected)
        flag = "OK  " if ok else "FAIL"
        print(f"  [{flag}] {label}: got {v!r}, expected {expected!r}")

    # Sheetnames preserved?
    print(f"\n--- TEMPLATE SHEET INTEGRITY ---")
    print(f"  original sheets preserved: {wb2.sheetnames}")

    # category_value 来自源表 ≠ 默认？检查源表的 product 1 category
    print(f"\n--- SOURCE ROW OVERRIDES ---")
    # Re-read source to see first product's actual values
    src_wb = openpyxl.load_workbook(mock_src, data_only=True, read_only=True)
    src_ws = src_wb.worksheets[0]
    src_rows = list(src_ws.iter_rows(values_only=True))
    src_hdrs = src_rows[0]
    def src_col(label):
        return src_hdrs.index(label)  # 0-based for tuple indexing
    print(f"  source product 1 name: {src_rows[1][src_col('产品名')][:60]}")
    print(f"  source product 1 category: {src_rows[1][src_col('产品类目')]}")
    print(f"  source product 1 brand: {src_rows[1][src_col('品牌')]}")
    print(f"  source variant 1: color={src_rows[1][src_col('规格1选项')]}, size={src_rows[1][src_col('规格2选项')]}")
    print(f"  source variant 1 sku: {src_rows[1][src_col('平台SKU')]}")
    print(f"  source variant 1 price: {src_rows[1][src_col('税前价格')]}  stock: {src_rows[1][src_col('库存')]}")
    print(f"  source variant 1 size_chart: {str(src_rows[1][src_col('尺码图')])[:60]}")
    print()
    print(f"  ⚠️  source category is 'Womenswear & Underwear>Women's Tops>Women's T-shirts'")
    print(f"     but PH default category_value is 'Men's Tops/T-shirts' (settings win)")
    print(f"     user must edit settings → category_value, or output will be mis-categorized")


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else None
    main(src)