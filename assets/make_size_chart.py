"""Generate a generic clothing size chart PNG.

Used as the default fallback for the TikTok Shop `size_chart` column when
the source EasyBoss table doesn't carry a size chart URL — TikTok's batch
upload validator flags empty size_chart cells as errors on clothing
products that have size variations (M / L / XL / …).

The PNG is shipped at `assets/default_size_chart.png` and is referenced
from the tool's default config via its public raw URL on GitHub.
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

OUT = Path(__file__).resolve().parent / "default_size_chart.png"

# Canvas
W, H = 1400, 800
BG = (255, 255, 255)
INK = (40, 40, 40)
GRID = (180, 180, 180)
HEAD_BG = (242, 244, 247)
HIGHLIGHT = (15, 122, 254)

# Default clothing measurements (cm, +/-1 cm tolerance).
# These are a generic Unisex T-Shirt chart — close enough for COD dropshipping
# where each listing has the same chart.
ROWS = [
    ("Size", "Length (cm)", "Chest (cm)", "Shoulder (cm)", "Sleeve (cm)", "Suggested Weight (kg)"),
    ("S",    "70", "96",  "44", "20", "45-55"),
    ("M",    "72", "100", "46", "21", "55-65"),
    ("L",    "74", "104", "48", "22", "65-75"),
    ("XL",   "76", "108", "50", "23", "75-85"),
    ("2XL",  "78", "112", "52", "24", "85-95"),
    ("3XL",  "80", "116", "54", "25", "95-110"),
]

NOTES = [
    "* All measurements are in centimetres (cm) and allow +/- 2 cm tolerance.",
    "* Please allow 1-3 cm deviation due to manual measurement.",
    "* Asian sizes tend to run smaller than EU/US — please check the chart carefully.",
    "* If you are between two sizes, we recommend ordering the larger one.",
]


def find_font(size: int) -> ImageFont.FreeTypeFont:
    candidates = [
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/msyh.ttc",
    ]
    for c in candidates:
        if Path(c).exists():
            try:
                return ImageFont.truetype(c, size)
            except OSError:
                continue
    return ImageFont.load_default()


def main():
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    title_font = find_font(40)
    head_font = find_font(26)
    cell_font = find_font(24)
    note_font = find_font(20)

    # ---- Title
    draw.text((W // 2, 50), "Size Chart - Unisex T-Shirt",
              fill=INK, font=title_font, anchor="mm")
    draw.text((W // 2, 95), "Please refer to this chart before ordering",
              fill=(110, 110, 110), font=note_font, anchor="mm")

    # ---- Table
    cols = len(ROWS[0])
    pad_x = 60
    pad_y = 150
    table_w = W - 2 * pad_x
    col_w = table_w // cols
    row_h = 60

    table_x0 = pad_x
    table_y0 = pad_y

    # Header row
    draw.rectangle(
        [table_x0, table_y0, table_x0 + table_w, table_y0 + row_h],
        fill=HEAD_BG,
        outline=GRID,
        width=2,
    )
    for c, val in enumerate(ROWS[0]):
        x = table_x0 + c * col_w + col_w // 2
        draw.text((x, table_y0 + row_h // 2), val,
                  fill=INK, font=head_font, anchor="mm")

    # Data rows
    for r, row in enumerate(ROWS[1:], 1):
        y = table_y0 + r * row_h
        if r % 2 == 1:
            draw.rectangle(
                [table_x0, y, table_x0 + table_w, y + row_h],
                fill=(248, 250, 252),
                outline=GRID,
                width=1,
            )
        else:
            draw.rectangle(
                [table_x0, y, table_x0 + table_w, y + row_h],
                fill=BG,
                outline=GRID,
                width=1,
            )
        for c, val in enumerate(row):
            x = table_x0 + c * col_w + col_w // 2
            fill = HIGHLIGHT if c == 0 else INK
            font = head_font if c == 0 else cell_font
            draw.text((x, y + row_h // 2), val,
                      fill=fill, font=font, anchor="mm")

    # ---- Outline
    draw.rectangle(
        [table_x0, table_y0, table_x0 + table_w, table_y0 + row_h * len(ROWS)],
        outline=INK,
        width=2,
    )

    # ---- Notes block
    notes_y = table_y0 + row_h * len(ROWS) + 50
    draw.text((pad_x, notes_y), "How to measure",
              fill=INK, font=head_font, anchor="lm")
    for i, line in enumerate(NOTES):
        draw.text((pad_x, notes_y + 40 + i * 32), line,
                  fill=(70, 70, 70), font=note_font, anchor="lm")

    img.save(OUT, "PNG", optimize=True)
    print(f"Wrote {OUT}  ({OUT.stat().st_size / 1024:.1f} KB)")


if __name__ == "__main__":
    main()
