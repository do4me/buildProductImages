#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build.py — directory walker & image generator
Python 3.12+

Behavior (per directory):
1) Require files:
   - background.png
   - background_bar_header.png
   - background_bar_footer.png
   - productNameAttr.txt (5 lines: font, bg_color, text_color, font_size, name_top_left x,y)
   - productInfoAttr.txt (4 lines: font, text_color, font_size, info_top_left x,y)
   - positions.txt (3 lines: centers for B, A, C as x,y)
   - maxheight.txt (3 lines: A_max_h, B_max_h, C_max_h; positive integers)
   If any missing → warn and skip directory.

2) Validate attr files, font existence, hex colors, positive font sizes.

3) Use Products.xlsx (row 2+):
   A=Folder, B=SKU, C=ProductName, D=ProductInfo, E=A rot, F=B rot, G=C rot.
   - Verify {SKU}_A.png / _B.png / _C.png exist in folder
   - Copy background.png → output/<Folder>/{SKU}_{ProductNameNoSpaces}.png
   - Compose: draw ProductName (with BG), draw ProductInfo; paste A/B/C centered at
     positions from positions.txt with the specified rotations (deg; negative=left, positive=right).

Scaling rule:
- Read A/B/C max heights from maxheight.txt (lines 1/2/3 respectively). Images are scaled
  down to those max heights (keeping aspect) BEFORE rotation.

Output:
- All generated files go under: <root>/output/<same directory structure as source>.
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Tuple
from contextlib import contextmanager

REQUIRED_FILES: tuple[str, ...] = (
    "background.png",
    "background_bar_header.png",
    "background_bar_footer.png",
    "productNameAttr.txt",
    "productInfoAttr.txt",
    "positions.txt",
    "maxheight.txt",
)

HEX_DIGITS = set("0123456789abcdefABCDEF")


@dataclass(slots=True)
class ProductInfoAttr:
    font_file: str
    text_color: str
    font_size: int  # in pt
    top_left: tuple[int, int]


@dataclass(slots=True)
class ProductNameAttr:
    font_file: str
    bg_color: str
    text_color: str
    font_size: int  # in pt
    top_left: tuple[int, int]


def _normalize_hex(s: str) -> str:
    s = s.strip()
    return s[1:] if s.startswith("#") else s


def _is_valid_hex_color(s: str) -> bool:
    s = _normalize_hex(s)
    return len(s) == 6 and all(c in HEX_DIGITS for c in s)


def _read_lines(path: Path) -> List[str]:
    # Robustly strip CR/LF and whitespace; keep UTF-8 BOM tolerant
    with path.open("r", encoding="utf-8-sig") as f:
        return [ln.rstrip("\r\n").strip() for ln in f.readlines()]


def _require_files(dir_path: Path) -> Tuple[bool, List[str]]:
    missing: List[str] = []
    for name in REQUIRED_FILES:
        if not (dir_path / name).is_file():
            missing.append(name)
    return (len(missing) == 0, missing)


def _font_exists(dir_path: Path, font_name: str) -> bool:
    p = Path(font_name)
    if not p.is_absolute():
        p = dir_path / p
    return p.is_file()


def _parse_xy(s: str, label: str) -> tuple[int, int]:
    try:
        x_str, y_str = [t.strip() for t in s.split(",")]
        return (int(x_str), int(y_str))
    except Exception:
        raise ValueError(f"Invalid {label} position: {s} (expected 'x,y')")


def _parse_product_info_attr(dir_path: Path) -> ProductInfoAttr:
    p = dir_path / "productInfoAttr.txt"
    lines = _read_lines(p)
    if len(lines) < 4:
        raise ValueError("productInfoAttr.txt must have 4 lines: font_file, text_color, font_size, top_left x,y")
    font_file, text_color, font_size_s, pos_s = lines[0], lines[1], lines[2], lines[3]

    if not _font_exists(dir_path, font_file):
        raise FileNotFoundError(f"Font file not found for productInfo: {font_file}")
    if not _is_valid_hex_color(text_color):
        raise ValueError(f"Invalid productInfo text color: {text_color}")
    try:
        font_size = int(font_size_s)
        if font_size <= 0:
            raise ValueError
    except Exception:
        raise ValueError(f"Invalid productInfo font size: {font_size_s}")

    pos = _parse_xy(pos_s, "productInfo")

    return ProductInfoAttr(font_file, "#" + _normalize_hex(text_color), font_size, pos)


def _parse_product_name_attr(dir_path: Path) -> ProductNameAttr:
    p = dir_path / "productNameAttr.txt"
    lines = _read_lines(p)
    if len(lines) < 5:
        raise ValueError("productNameAttr.txt must have 5 lines: font_file, bg_color, text_color, font_size, top_left x,y")
    font_file, bg_color, text_color, font_size_s, pos_s = lines[0], lines[1], lines[2], lines[3], lines[4]

    if not _font_exists(dir_path, font_file):
        raise FileNotFoundError(f"Font file not found for productName: {font_file}")
    if not _is_valid_hex_color(bg_color):
        raise ValueError(f"Invalid productName background color: {bg_color}")
    if not _is_valid_hex_color(text_color):
        raise ValueError(f"Invalid productName text color: {text_color}")
    try:
        font_size = int(font_size_s)
        if font_size <= 0:
            raise ValueError
    except Exception:
        raise ValueError(f"Invalid productName font size: {font_size_s}")

    pos = _parse_xy(pos_s, "productName")

    return ProductNameAttr(
        font_file,
        "#" + _normalize_hex(bg_color),
        "#" + _normalize_hex(text_color),
        font_size,
        pos,
    )


def validate_directory(dir_path: Path) -> tuple[bool, str | None, ProductNameAttr | None, ProductInfoAttr | None]:
    ok, missing = _require_files(dir_path)
    if not ok:
        return False, f"Missing required files: {', '.join(missing)}", None, None
    try:
        info_attr = _parse_product_info_attr(dir_path)
        name_attr = _parse_product_name_attr(dir_path)
    except Exception as e:
        return False, str(e), None, None
    return True, None, name_attr, info_attr

# ---- Excel parsing, positions, & generation ----
try:
    import openpyxl  # type: ignore
except Exception:
    openpyxl = None

try:
    from PIL import Image, ImageDraw, ImageFont  # type: ignore
except Exception:
    Image = None  # type: ignore
    ImageDraw = None  # type: ignore
    ImageFont = None  # type: ignore


@dataclass(slots=True)
class ExcelRow:
    folder: str
    sku: str
    product_name: str
    product_info: str
    rot_a: int
    rot_b: int
    rot_c: int


def read_products_excel(xlsx_path: Path) -> list[ExcelRow]:
    if openpyxl is None:
        raise RuntimeError("openpyxl is required to read Products.xlsx. pip install openpyxl")
    if not xlsx_path.is_file():
        raise FileNotFoundError(f"Products.xlsx not found: {xlsx_path}")

    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb.active

    def _to_int(v):
        try:
            return int(v)
        except Exception:
            return 0

    rows: list[ExcelRow] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None:
            continue
        folder = (row[0] or "").strip()
        sku = (row[1] or "").strip()
        product_name = (row[2] or "").strip()
        product_info = (row[3] or "").strip()
        rot_a = _to_int(row[4])
        rot_b = _to_int(row[5])
        rot_c = _to_int(row[6])
        if not folder or not sku:
            continue
        rows.append(ExcelRow(folder, sku, product_name, product_info, rot_a, rot_b, rot_c))
    wb.close()
    return rows


def check_sku_images(dir_path: Path, sku: str) -> tuple[bool, list[str]]:
    expected = [f"{sku}_A.png", f"{sku}_B.png", f"{sku}_C.png"]
    missing: list[str] = [fn for fn in expected if not (dir_path / fn).is_file()]
    return (len(missing) == 0, missing)


def copy_background_to_out(dir_path: Path, out_dir: Path, sku: str, product_name: str) -> Path:
    """
    Copy background.png from dir_path to out_dir as {sku}_{ProductNameNoSpaces}.png.
    Ensures out_dir exists. Returns destination path.
    """
    if Image is None:
        raise RuntimeError("Pillow (PIL) is required. pip install pillow")
    src = dir_path / "background.png"
    if not src.is_file():
        raise FileNotFoundError(f"background.png not found in {dir_path}")
    out_dir.mkdir(parents=True, exist_ok=True)
    safe_name = sku + "_" + product_name.replace(" ", "")
    dst = out_dir / f"{safe_name}.png"
    with Image.open(src) as im:
        im.copy().save(dst)
    return dst


def read_positions(dir_path: Path) -> tuple[tuple[int, int], tuple[int, int], tuple[int, int]]:
    p = dir_path / "positions.txt"
    lines = _read_lines(p)
    if len(lines) < 3:
        raise ValueError("positions.txt must have 3 lines for B, A, C centers (x,y)")
    b = _parse_xy(lines[0], "B center")
    a = _parse_xy(lines[1], "A center")
    c = _parse_xy(lines[2], "C center")
    return b, a, c


def read_max_heights(dir_path: Path) -> tuple[int, int, int]:
    """
    Read maxheight.txt: line1=A max height, line2=B, line3=C (positive integers).
    """
    p = dir_path / "maxheight.txt"
    lines = _read_lines(p)
    if len(lines) < 3:
        raise ValueError("maxheight.txt must have 3 lines: A_max_h, B_max_h, C_max_h (positive integers)")
    def _parse_pos_int(s: str, label: str) -> int:
        try:
            v = int(s)
            if v <= 0:
                raise ValueError
            return v
        except Exception:
            raise ValueError(f"Invalid {label} in maxheight.txt: '{s}' (must be positive integer)")
    a_h = _parse_pos_int(lines[0], "A max height")
    b_h = _parse_pos_int(lines[1], "B max height")
    c_h = _parse_pos_int(lines[2], "C max height")
    return a_h, b_h, c_h


def _paste_centered(base: Image.Image, overlay: Image.Image, center: tuple[int, int]) -> None:
    x, y = center
    w, h = overlay.size
    pos = (int(x - w / 2), int(y - h / 2))
    # Use alpha channel as mask if present
    mask = overlay.split()[3] if overlay.mode == "RGBA" and len(overlay.getbands()) == 4 else None
    base.paste(overlay, pos, mask)


@contextmanager
def open_scale_rotate(path: Path, deg: int, max_h: int):
    """
    Context manager that opens a PNG, scales down to max_h (if taller, keep aspect), then rotates.
    Yields an RGBA Image and guarantees close().
    """
    img = Image.open(path).convert("RGBA")
    try:
        if max_h and img.height > max_h:
            scale = max_h / img.height
            new_w = int(round(img.width * scale))
            img = img.resize((new_w, max_h), resample=Image.LANCZOS)
        if deg:
            img = img.rotate(-deg, expand=True, resample=Image.BICUBIC)
        yield img
    finally:
        try:
            img.close()
        except Exception:
            pass


def generate_images(
    root: Path,
    output_root: Path,
    dir_path: Path,
    name_attr: ProductNameAttr,
    info_attr: ProductInfoAttr,
    row: ExcelRow,
) -> None:
    ok, missing = check_sku_images(dir_path, row.sku)
    if not ok:
        print(f"[SKIP] {dir_path.name}/{row.sku}: missing {', '.join(missing)}")
        return

    # Compute the mirrored output directory under output_root, preserving structure
    rel_dir = dir_path.relative_to(root)
    out_dir = output_root / rel_dir

    out_path = copy_background_to_out(dir_path, out_dir, row.sku, row.product_name)

    with Image.open(out_path).convert("RGBA") as base:
        draw = ImageDraw.Draw(base)
        try:
            name_font = ImageFont.truetype(str(dir_path / name_attr.font_file), name_attr.font_size)
            info_font = ImageFont.truetype(str(dir_path / info_attr.font_file), info_attr.font_size)
        except Exception as e:
            print(f"[SKIP] {dir_path.name}/{row.sku}: cannot load fonts: {e}")
            return

        # ProductName with background (fixed 120px height + header/footer images)
        name_text = row.product_name
        bbox = draw.textbbox((0, 0), name_text, font=name_font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        nx, ny = name_attr.top_left
        BG_H = 119  # px

        # Load header/footer images from source dir
        try:
            header_img = Image.open(dir_path / "background_bar_header.png").convert("RGBA")
            footer_img = Image.open(dir_path / "background_bar_footer.png").convert("RGBA")
        except Exception as e:
            print(f"[SKIP] {dir_path.name}/{row.sku}: cannot open header/footer: {e}")
            return

        # Use header/footer as-is (assumed already 120px height)
        header_resized = header_img
        footer_resized = footer_img

        # Draw background behind text area only
        draw.rectangle(
            [nx + header_resized.width, ny, nx + header_resized.width + tw, ny + BG_H],
            fill=name_attr.bg_color,
        )

        # Paste header and footer
        base.alpha_composite(header_resized, (nx, ny))
        base.alpha_composite(footer_resized, (nx + header_resized.width + tw, ny))

        # Vertically center the text within the 120px bar
        ty = ny + (BG_H - th) // 3
        tx = nx + header_resized.width
        draw.text((tx, ty), name_text, font=name_font, fill=name_attr.text_color)

        # ProductInfo (text only)
        draw.text(info_attr.top_left, row.product_info, font=info_font, fill=info_attr.text_color)

        # Read centers and max heights
        try:
            pos_b, pos_a, pos_c = read_positions(dir_path)
        except Exception as e:
            print(f"[SKIP] {dir_path.name}/{row.sku}: {e}")
            return

        try:
            max_a, max_b, max_c = read_max_heights(dir_path)
        except Exception as e:
            print(f"[SKIP] {dir_path.name}/{row.sku}: {e}")
            return

        # Apply scaling & rotation, then paste centered
        with open_scale_rotate(dir_path / f"{row.sku}_B.png", row.rot_b, max_b) as img_b:
            _paste_centered(base, img_b, pos_b)

        with open_scale_rotate(dir_path / f"{row.sku}_A.png", row.rot_a, max_a) as img_a:
            _paste_centered(base, img_a, pos_a)

        with open_scale_rotate(dir_path / f"{row.sku}_C.png", row.rot_c, max_c) as img_c:
            _paste_centered(base, img_c, pos_c)

        base.save(out_path)
        rel_saved = out_path.relative_to(output_root)
        print(f"[OK] Composed image saved: {rel_saved}")


# ---- Runner ----

def iter_directories(root: Path) -> Iterable[Path]:
    for entry in sorted(root.iterdir()):
        if entry.is_dir() and not entry.name.startswith("."):
            yield entry


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate directories and generate images.")
    parser.add_argument("root", nargs="?", default=Path.cwd(), type=Path, help="Root folder (default: cwd)")
    parser.add_argument("--xlsx", default="Products.xlsx", help="Excel file next to build.py")
    parser.add_argument("--out", default="output", help="Output directory (default: 'output' under root)")
    args = parser.parse_args(argv)

    root: Path = args.root.resolve()
    if not root.is_dir():
        print(f"Error: root is not a directory: {root}", file=sys.stderr)
        return 2

    output_root: Path = (root / args.out).resolve()
    output_root.mkdir(parents=True, exist_ok=True)

    try:
        rows = read_products_excel((root / args.xlsx).resolve())
    except Exception as e:
        print(f"Error reading Excel: {e}", file=sys.stderr)
        return 2

    print(f"Scanning: {root}  (rows loaded: {len(rows)})")
    print(f"Output root: {output_root}")

    by_folder: dict[str, list[ExcelRow]] = {}
    for r in rows:
        by_folder.setdefault(r.folder, []).append(r)

    any_processed = False
    for d in iter_directories(root):
        wanted = by_folder.get(d.name)
        if not wanted:
            continue
        print(f"\n── Checking directory: {d.name} ──")
        valid, err, name_attr, info_attr = validate_directory(d)
        if not valid or not (name_attr and info_attr):
            print(f"[SKIP] {d.name}: {err}")
            continue
        for row in wanted:
            print(
                f"  → SKU={row.sku}  Name='{row.product_name}'  Info='{row.product_info}'  "
                f"Rot(A,B,C)=({row.rot_a},{row.rot_b},{row.rot_c})"
            )
            any_processed = True
            generate_images(root, output_root, d, name_attr, info_attr, row)

    if not any_processed:
        print("No valid rows processed.")
        return 1

    print("\nDone.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
