"""Core conversion orchestration: source → Products → TikTok rows → xlsx.

This module also handles an optional image download pass: many EasyBoss
export tables reference Shopee-hosted image URLs that expire. TikTok Shop
requires a public, persistent image URL, so we offer a helper that downloads
each unique image into a sibling `images/` folder next to the output xlsx
and writes an `_image_manifest.txt` for the user to upload to TikTok Media
Center.
"""
from __future__ import annotations

import hashlib
import random
import re
import string
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from .source_reader import Product, read_source
from .tiktok_writer import build_rows_for_product, write_tiktok_xlsx


ProgressFn = Callable[[float, str], None]   # (0.0..1.0, status_text)


def generate_random_suffix(length: int = 3, lowercase: bool = False) -> str:
    """Generate a short random token to differentiate title copies (防查重).

    v1.2 used 10-char lowercase; we honor that default.
    """
    length = max(1, min(int(length or 3), 16))
    chars = (
        string.ascii_lowercase + string.digits
        if lowercase
        else string.ascii_uppercase + string.digits
    )
    safe_first = (
        "abcdefghjkmnpqrstuvwxyz23456789"
        if lowercase
        else "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    )
    first = random.choice(safe_first)
    rest = "".join(random.choice(chars) for _ in range(length - 1))
    return first + rest


def generate_copy_suffixes(copies: int, length: int) -> list[str]:
    """Return ['', s1, s2, ..., s_{copies-1}] — first copy has no suffix."""
    copies = max(1, int(copies))
    length = max(1, int(length or 10))
    if copies <= 1:
        return [""]
    out = [""]
    seen: set[str] = set()
    while len(out) < copies:
        s = generate_random_suffix(length, lowercase=True)
        if s in seen:
            continue
        seen.add(s)
        out.append(s)
    return out


@dataclass
class ConvertResult:
    output_path: Path
    product_count: int
    row_count: int
    image_dir: Path | None = None
    image_manifest: Path | None = None
    image_download_count: int = 0
    image_failed: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Image download helper
# ---------------------------------------------------------------------------


def _filename_from_url(url: str, idx: int, ext_default: str = ".jpg") -> str:
    """Derive a safe local filename from a remote image URL."""
    try:
        path = urllib.parse.urlparse(url).path
    except Exception:
        path = url
    last = path.rsplit("/", 1)[-1] if path else f"image_{idx}"
    # Strip query string artifacts
    last = last.split("?", 1)[0]
    # Keep only safe characters
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", last)
    if not safe or safe.startswith("."):
        safe = f"image_{idx}{ext_default}"
    if "." not in safe:
        safe = safe + ext_default
    return safe[:120]


def _safe_path_for(url: str, idx: int) -> str:
    """Make a unique filename by hashing the URL to avoid collisions."""
    h = hashlib.md5(url.encode("utf-8", "replace")).hexdigest()[:8]
    base = _filename_from_url(url, idx)
    stem, dot, ext = base.rpartition(".")
    if not dot:
        return f"{base}_{h}"
    return f"{stem}_{h}.{ext}"


def download_images(
    products: list[Product],
    output_dir: Path,
    progress: ProgressFn | None = None,
    timeout: float = 8.0,
) -> tuple[Path, Path, int, list[str]]:
    """Download all unique product images into `<output_dir>/images/`.

    Returns: (image_dir, manifest_path, success_count, failed_urls)

    The manifest is a plain-text list (one URL → local file per line) so the
    user can pipe it into their TikTok Media Center upload flow.
    """
    # Collect all unique image URLs
    urls: list[str] = []
    seen: set[str] = set()
    for p in products:
        for img in [p.size_chart, *p.images, *(v.sku_image for v in p.variants)]:
            if img and img not in seen:
                seen.add(img)
                urls.append(img)

    image_dir = output_dir / "images"
    image_dir.mkdir(parents=True, exist_ok=True)
    failed: list[str] = []
    success = 0

    if progress:
        progress(0.7, f"正在下载 {len(urls)} 张图片（最多 {int(timeout)}s/张）…")

    for idx, url in enumerate(urls, 1):
        fname = _safe_path_for(url, idx)
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
            })
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = resp.read(5_000_000)  # cap at 5MB
            (image_dir / fname).write_bytes(data)
            success += 1
        except Exception:
            failed.append(url)
        if progress and urls:
            progress(0.7 + 0.1 * (idx / len(urls)), f"已处理 {idx}/{len(urls)} 张图片…")

    manifest = output_dir / "_image_manifest.txt"
    lines = ["# TikTok PH Converter - image manifest",
            "# Format: <remote_url>  ->  <local_file>",
            "# After uploading these to TikTok Media Center, replace each",
            "# remote URL in the output xlsx with the TikTok-hosted URL.",
            ""]
    for idx, url in enumerate(urls, 1):
        local = _safe_path_for(url, idx)
        ok = "OK " if url not in failed else "ERR"
        lines.append(f"[{ok}] {url}  ->  images/{local}")
    manifest.write_text("\n".join(lines), encoding="utf-8")

    if progress:
        progress(0.8, f"图片下载完成：{success}/{len(urls)} 成功")
    return image_dir, manifest, success, failed


def convert_source(
    source_xlsx: Path,
    output_dir: Path,
    settings: dict[str, Any],
    template_src: Path,
    progress: ProgressFn | None = None,
    column_mapping: dict[str, str] | None = None,
    download_imgs: bool = True,
) -> ConvertResult:
    """Read source, apply settings, write TikTok batch upload xlsx.

    Output filename: <source-stem>_TKPH_<timestamp>.xlsx in `output_dir`.

    ROW STRATEGY: each (color, size) variant in the source becomes one row.
    `output_copies` duplicates each variant N times with random title
    suffixes to avoid TikTok's duplicate-listing detection (防查重).

    If `download_imgs=True` (default), every remote image referenced in the
    source is downloaded to `<output_dir>/images/` and a manifest is
    written next to the output xlsx for the user to re-host via TikTok
    Media Center (Shopee cf.shopee.ph URLs expire and won't load from
    TikTok).
    """
    if progress:
        progress(0.02, "正在读取源表格…")
    products = read_source(source_xlsx, column_mapping=column_mapping)
    if not products:
        raise ValueError("源表格中没有可识别的产品数据。")

    copies = max(1, int(settings.get("output_copies", 1) or 1))
    use_suffix = bool(settings.get("title_random_suffix_enabled", False))
    suffix_len = int(settings.get("random_suffix_length", 10) or 10)

    if progress:
        progress(0.25, f"已识别 {len(products)} 个产品（{sum(len(p.variants) for p in products)} 个变体），正在生成行…")
    rows: list[dict[str, Any]] = []
    total = len(products)
    for i, p in enumerate(products, 1):
        if use_suffix and copies > 1:
            suffixes = generate_copy_suffixes(copies, suffix_len)
        else:
            suffixes = [""] * copies
        rows.extend(build_rows_for_product(p, settings, copy_suffixes=suffixes))
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

    image_dir = image_manifest = None
    img_count = 0
    img_failed: list[str] = []
    if download_imgs:
        image_dir, image_manifest, img_count, img_failed = download_images(
            products, output_dir, progress=progress,
        )

    if progress:
        msg = f"完成。共 {len(products)} 个产品 / {len(rows)} 行 → {out_path.name}"
        if img_count or img_failed:
            msg += f"  (图片: {img_count} 成功"
            if img_failed:
                msg += f", {len(img_failed)} 失败"
            msg += ")"
        progress(1.0, msg)
    return ConvertResult(
        output_path=out_path,
        product_count=len(products),
        row_count=len(rows),
        image_dir=image_dir,
        image_manifest=image_manifest,
        image_download_count=img_count,
        image_failed=img_failed,
    )