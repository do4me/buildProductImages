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
    font_size: int
    top_left: tuple[int, int]


@dataclass(slots=True)
class ProductNameAttr:
    font_file: str
    bg_color: str
    text_color: str
    font_size: int
    top_left: tuple[int, int]


def _normalize_hex(s: str) -> str:
    s = s.strip()
    return s[1:] if s.startswith("#") else s


def _is_valid_hex_color(s: str) -> bool:
    s = _normalize_hex(s)
    return len(s) == 6 and all(c in HEX_DIGITS for c in s)


def _read_lines(path: Path) -> List[str]:
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
        raise ValueError("productInfoAttr.txt must have 4 lines")
    font_file, text_color, font_size_s, pos_s = lines

    if not _font_exists(dir_path, font_file):
        raise FileNotFoundError(f"Font file not found for productInfo: {font_file}")
    if not _is_valid_hex_color(text_color):
        raise ValueError(f"Invalid productInfo text color: {text_color}")
    font_size = int(font_size_s)
    if font_size <= 0:
        raise ValueError("Invalid productInfo font size")

    pos = _parse_xy(pos_s, "productInfo")
    return ProductInfoAttr(font_file, "#" + _normalize_hex(text_color), font_size, pos)


def _parse_product_name_attr(dir_path: Path) -> ProductNameAttr:
    p = dir_path / "productNameAttr.txt"
    lines = _read_lines(p)
    if len(lines) < 5:
        raise ValueError("productNameAttr.txt must have 5 lines")
    font_file, bg_color, text_color, font_size_s, pos_s = lines

    if not _font_exists(dir_path, font_file):
        raise FileNotFoundError(f"Font file not found for productName: {font_file}")
    if not _is_valid_hex_color(bg_color):
        raise ValueError(f"Invalid productName background color: {bg_color}")
    if not _is_valid_hex_color(text_color):
        raise ValueError(f"Invalid productName text color: {text_color}")
    font_size = int(font_size_s)
    if font_size <= 0:
        raise ValueError("Invalid productName font size")

    pos = _parse_xy(pos_s, "productName")
    return ProductNameAttr(
        font_file,
        "#" + _normalize_hex(bg_color),
        "#" + _normalize_hex(text_color),
        font_size,
        pos,
    )


def validate_directory(dir_path: Path):
    ok, missing = _require_files(dir_path)
    if not ok:
        return False, f"Missing required files: {', '.join(missing)}", None, None
    try:
        info_attr = _parse_product_info_attr(dir_path)
        name_attr = _parse_product_name_attr(dir_path)
    except Exception as e:
        return False, str(e), None, None
    return True, None, name_attr, info_attr


try:
    import openpyxl
except Exception:
    openpyxl = None

try:
    from PIL import Image, ImageDraw, ImageFont
except Exception:
    Image = None
    ImageDraw = None
    ImageFont = None


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
        raise RuntimeError("openpyxl required")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb.active
    rows: list[ExcelRow] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        folder, sku, name, info, ra, rb, rc = row[:7]
        if not folder or not sku:
            continue
        rows.append(ExcelRow(folder.strip(), sku.strip(), (name or "").strip(),
                             (info or "").strip(), int(ra or 0), int(rb or 0), int(rc or 0)))
    wb.close()
    return rows


def check_sku_images(dir_path: Path, sku: str) -> tuple[bool, list[str]]:
    expected = [f"{sku}_A.png", f"{sku}_B.png", f"{sku}_C.png"]
    missing = [fn for fn in expected if not (dir_path / fn).is_file()]
    return (len(missing) == 0, missing)


def copy_background_to_out(dir_path: Path, out_dir: Path, sku: str, product_name: str) -> Path:
    src = dir_path / "background.png"
    out_dir.mkdir(parents=True, exist_ok=True)
    safe_name = sku + "_" + product_name.replace(" ", "")
    dst = out_dir / f"{safe_name}.png"
    with Image.open(src) as im:
        im.copy().save(dst)
    return dst


def read_positions(dir_path: Path):
    lines = _read_lines(dir_path / "positions.txt")
    b = _parse_xy(lines[0], "B center")
    a = _parse_xy(lines[1], "A center")
    c = _parse_xy(lines[2], "C center")
    return b, a, c


def read_max_heights(dir_path: Path):
    lines = _read_lines(dir_path / "maxheight.txt")
    vals = []
    for idx, label in enumerate(("A", "B", "C")):
        v = int(lines[idx])
        if v <= 0:
            raise ValueError(f"Invalid {label} max height")
        vals.append(v)
    return tuple(vals)


def _paste_centered(base: Image.Image, overlay: Image.Image, center: tuple[int, int]) -> None:
    x, y = center
    w, h = overlay.size
    pos = (int(x - w / 2), int(y - h / 2))
    mask = overlay.split()[3] if overlay.mode == "RGBA" else None
    base.paste(overlay, pos, mask)


@contextmanager
def open_scale_rotate(path: Path, deg: int, max_h: int):
    img = Image.open(path).convert("RGBA")
    try:
        if max_h and img.height > max_h:
            scale = max_h / img.height
            img = img.resize((int(round(img.width * scale)), max_h), resample=Image.LANCZOS)
        if deg:
            img = img.rotate(-deg, expand=True, resample=Image.BICUBIC)
        yield img
    finally:
        img.close()


def generate_images(root: Path, output_root: Path, dir_path: Path,
                    name_attr: ProductNameAttr, info_attr: ProductInfoAttr, row: ExcelRow) -> None:
    ok, missing = check_sku_images(dir_path, row.sku)
    if not ok:
        print(f"[SKIP] {row.sku}: missing {', '.join(missing)}")
        return

    rel_dir = dir_path.relative_to(root)
    out_dir = output_root / rel_dir
    out_path = copy_background_to_out(dir_path, out_dir, row.sku, row.product_name)

    with Image.open(out_path).convert("RGBA") as base:
        draw = ImageDraw.Draw(base)
        try:
            name_font = ImageFont.truetype(str(dir_path / name_attr.font_file), name_attr.font_size)
            info_font = ImageFont.truetype(str(dir_path / info_attr.font_file), info_attr.font_size)
        except Exception as e:
            print(f"[SKIP] {row.sku}: cannot load fonts: {e}")
            return

        # ==== 底部对齐的 ProductName 绘制 ====
        name_text = row.product_name
        bbox = draw.textbbox((0, 0), name_text, font=name_font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        nx, ny = name_attr.top_left
        BG_H = 120

        try:
            header_img = Image.open(dir_path / "background_bar_header.png").convert("RGBA")
            footer_img = Image.open(dir_path / "background_bar_footer.png").convert("RGBA")
        except Exception as e:
            print(f"[SKIP] {row.sku}: cannot open header/footer: {e}")
            return

        header_w, header_h = header_img.size
        footer_w, footer_h = footer_img.size
        bar_bottom = ny + BG_H

        # 背景矩形
        draw.rectangle(
            [nx + header_w-1, bar_bottom - BG_H, nx + header_w + tw, bar_bottom-1],
            fill=name_attr.bg_color,
        )
        # 底部对齐贴 header/footer
        base.alpha_composite(header_img, (nx, bar_bottom - header_h))
        base.alpha_composite(footer_img, (nx + header_w + tw, bar_bottom - footer_h))

        # 底部对齐绘制文字
        tx = nx + header_w
        ty = bar_bottom - 100
        draw.text((tx, ty), name_text, font=name_font, fill=name_attr.text_color)

        # ProductInfo
        draw.text(info_attr.top_left, row.product_info, font=info_font, fill=info_attr.text_color)

        # 图片贴合
        try:
            pos_b, pos_a, pos_c = read_positions(dir_path)
            max_a, max_b, max_c = read_max_heights(dir_path)
        except Exception as e:
            print(f"[SKIP] {row.sku}: {e}")
            return

        with open_scale_rotate(dir_path / f"{row.sku}_B.png", row.rot_b, max_b) as img_b:
            _paste_centered(base, img_b, pos_b)
        with open_scale_rotate(dir_path / f"{row.sku}_A.png", row.rot_a, max_a) as img_a:
            _paste_centered(base, img_a, pos_a)
        with open_scale_rotate(dir_path / f"{row.sku}_C.png", row.rot_c, max_c) as img_c:
            _paste_centered(base, img_c, pos_c)

        base.save(out_path)
        print(f"[OK] {out_path.relative_to(output_root)}")


def iter_directories(root: Path) -> Iterable[Path]:
    for entry in sorted(root.iterdir()):
        if entry.is_dir() and not entry.name.startswith("."):
            yield entry


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("root", nargs="?", default=Path.cwd(), type=Path)
    parser.add_argument("--xlsx", default="Products.xlsx")
    parser.add_argument("--out", default="output")
    args = parser.parse_args(argv)

    root = args.root.resolve()
    output_root = (root / args.out).resolve()
    output_root.mkdir(parents=True, exist_ok=True)

    try:
        rows = read_products_excel(root / args.xlsx)
    except Exception as e:
        print(f"Error reading Excel: {e}", file=sys.stderr)
        return 2

    by_folder: dict[str, list[ExcelRow]] = {}
    for r in rows:
        by_folder.setdefault(r.folder, []).append(r)

    any_processed = False
    for d in iter_directories(root):
        wanted = by_folder.get(d.name)
        if not wanted:
            continue
        valid, err, name_attr, info_attr = validate_directory(d)
        if not valid:
            print(f"[SKIP] {d.name}: {err}")
            continue
        for row in wanted:
            any_processed = True
            generate_images(root, output_root, d, name_attr, info_attr, row)

    if not any_processed:
        print("No valid rows processed.")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())