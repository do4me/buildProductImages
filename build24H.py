#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
build24H.py
Python 3.12.3

Behavior:
- Traverse immediate subdirectories of the script's working directory.
- For each directory:
  1) Ensure background_24H.png and background_24H_Scissors.png exist; otherwise log and skip.
  2) Ensure positions_24H.txt exists with EXACTLY 6 lines of "x,y" center coordinates:
        1: avatar center
        2: {SKU}_M1.png center
        3: {SKU}_M2.png center
        4: background_24H_Scissors.png center
        5: {SKU}_A.png center
        6: {SKU}_B.png center
     If invalid, log and skip.
  3) Ensure maxheight_24H.txt exists with EXACTLY 6 positive numbers (target heights for the same six layers).
  4) Ensure rotate_24H.txt exists with EXACTLY 6 numbers (degrees). Positive = clockwise, negative = counter-clockwise.
  5) Read Products.xlsx (in same dir as this script) with columns:
        Folder, SKU, Avatar, ProductName
     Process rows where Folder == current directory name.
     Validate required files: {SKU}_A.png, {SKU}_B.png, {SKU}_M1.png, {SKU}_M2.png, {Avatar}.png
  6) Create output by copying background_24H.png as a base canvas in memory,
     paste layers at their centers (order: avatar, M1, M2, A, B), then overlay background_24H_Scissors on top.
     Each layer is first scaled to target height, then rotated by its angle from rotate_24H.txt.
     Save as: {SKU}_{ProductNameNoSpaces}_24H.png
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Tuple, List, Dict

import pandas as pd
from PIL import Image

REQUIRED_BG = ["background_24H.png", "background_24H_Scissors.png"]
POSITIONS_FILE = "positions_24H.txt"
MAXHEIGHTS_FILE = "maxheight_24H.txt"
ROTATE_FILE = "rotate_24H.txt"
EXCEL_NAME = "Products.xlsx"

# Index mapping (1-based in spec -> 0-based here)
IDX_AVATAR = 0
IDX_M1 = 1
IDX_M2 = 2
IDX_SCISSORS = 3
IDX_A = 4
IDX_B = 5

def parse_center(line: str) -> Tuple[int, int]:
    parts = [p.strip() for p in line.split(",")]
    if len(parts) != 2:
        raise ValueError(f"Invalid coord '{line}'. Must be 'x,y'")
    x = int(round(float(parts[0])))
    y = int(round(float(parts[1])))
    return x, y

def load_positions_file(p: Path) -> List[Tuple[int, int]]:
    text = p.read_text(encoding="utf-8").strip().splitlines()
    if len(text) != 6:
        raise ValueError(f"{p.name} must contain exactly 6 lines, got {len(text)}")
    return [parse_center(line) for line in text]

def parse_height(line: str) -> int:
    val = int(round(float(line.strip())))
    if val <= 0:
        raise ValueError(f"height must be > 0, got {val}")
    return val

def load_maxheights_file(p: Path) -> List[int]:
    text = p.read_text(encoding="utf-8").strip().splitlines()
    if len(text) != 6:
        raise ValueError(f"{p.name} must contain exactly 6 lines, got {len(text)}")
    return [parse_height(line) for line in text]

def parse_angle(line: str) -> float:
    # Accept floats; positive = clockwise, negative = counter-clockwise (per spec).
    # PIL rotate(angle) uses positive = counter-clockwise, so we will invert sign at rotate time.
    try:
        return float(line.strip())
    except Exception:
        raise ValueError(f"invalid angle '{line}'")

def load_rotate_file(p: Path) -> List[float]:
    text = p.read_text(encoding="utf-8").strip().splitlines()
    if len(text) != 6:
        raise ValueError(f"{p.name} must contain exactly 6 lines, got {len(text)}")
    return [parse_angle(line) for line in text]

def ensure_png_name(name_from_excel: str) -> str:
    s = (name_from_excel or "").strip()
    if not s:
        return s
    return s if s.lower().endswith(".png") else f"{s}.png"

def center_paste(base: Image.Image, overlay: Image.Image, center: Tuple[int, int]) -> None:
    ox, oy = center
    w, h = overlay.size
    x = int(round(ox - w / 2))
    y = int(round(oy - h / 2))
    base.alpha_composite(overlay, dest=(x, y))

def scale_to_height(img: Image.Image, target_h: int) -> Image.Image:
    if target_h <= 0:
        raise ValueError(f"Invalid target height {target_h}")
    w, h = img.size
    if h == target_h:
        return img
    new_w = max(1, int(round(w * (target_h / h))))
    return img.resize((new_w, target_h), Image.LANCZOS)

def rotate_image(img: Image.Image, angle_deg_spec: float) -> Image.Image:
    # Spec: positive = clockwise; PIL: positive = counter-clockwise.
    pil_angle = -angle_deg_spec
    # Use bicubic, expand to keep full bounds. For RGBA, transparency is preserved.
    return img.rotate(pil_angle, resample=Image.BICUBIC, expand=True)

def validate_required_files(folder: Path) -> Tuple[Path, Path, Path, Path, Path]:
    bg = folder / "background_24H.png"
    fg = folder / "background_24H_Scissors.png"
    pos = folder / POSITIONS_FILE
    mh = folder / MAXHEIGHTS_FILE
    rot = folder / ROTATE_FILE
    missing = [p.name for p in [bg, fg, pos, mh, rot] if not p.exists()]
    if missing:
        raise FileNotFoundError(f"Missing {', '.join(missing)}")
    return bg, fg, pos, mh, rot

def read_products(excel_path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        raise RuntimeError(f"Failed to read {excel_path.name}: {e}")
    expected_cols = {"Folder", "SKU", "Avatar", "ProductName"}
    missing = expected_cols - set(df.columns)
    if missing:
        raise RuntimeError(f"{excel_path.name} missing columns: {', '.join(sorted(missing))}")
    for col in ["Folder", "SKU", "Avatar", "ProductName"]:
        df[col] = df[col].astype(str).fillna("").str.strip()
    return df

def process_folder(folder: Path, rows: pd.DataFrame) -> None:
    # Step 1: required files
    try:
        bg_path, scissors_path, pos_path, mh_path, rot_path = validate_required_files(folder)
    except Exception as e:
        print(f"[SKIP] {folder.name}: {e}")
        return

    # Step 2: parse positions, target heights, and rotation angles
    try:
        centers = load_positions_file(pos_path)
    except Exception as e:
        print(f"[SKIP] {folder.name}: invalid {POSITIONS_FILE}: {e}")
        return
    try:
        heights = load_maxheights_file(mh_path)
    except Exception as e:
        print(f"[SKIP] {folder.name}: invalid {MAXHEIGHTS_FILE}: {e}")
        return
    try:
        angles = load_rotate_file(rot_path)
    except Exception as e:
        print(f"[SKIP] {folder.name}: invalid {ROTATE_FILE}: {e}")
        return

    if rows.empty:
        print(f"[INFO] {folder.name}: no matching rows in {EXCEL_NAME}")
        return

    # Preload backgrounds
    try:
        bg_src = Image.open(bg_path).convert("RGBA")
        scissors_img_orig = Image.open(scissors_path).convert("RGBA")
    except Exception as e:
        print(f"[SKIP] {folder.name}: failed to open background(s): {e}")
        return

    # Pre-scale and pre-rotate the scissors overlay once (height line 4, angle line 4)
    try:
        scissors_img_scaled = scale_to_height(scissors_img_orig, heights[IDX_SCISSORS])
        scissors_img_final = rotate_image(scissors_img_scaled, angles[IDX_SCISSORS])
    except Exception as e:
        print(f"[SKIP] {folder.name}: prepare scissors failed: {e}")
        return

    # Process each product row
    for i, row in rows.iterrows():
        sku = row["SKU"].strip()
        avatar_name = ensure_png_name(row["Avatar"])
        product_name = row["ProductName"].strip()
        if not sku:
            print(f"[WARN] {folder.name}: row {i+2} empty SKU — skipping")
            continue
        if not avatar_name:
            print(f"[WARN] {folder.name}: row {i+2} empty Avatar — skipping")
            continue
        if not product_name:
            print(f"[WARN] {folder.name}: row {i+2} empty ProductName — skipping")
            continue

        a_path = folder / f"{sku}_A.png"
        b_path = folder / f"{sku}_B.png"
        m1_path = folder / f"{sku}_M1.png"
        m2_path = folder / f"{sku}_M2.png"
        avatar_path = folder / avatar_name

        missing_assets = [p.name for p in [a_path, b_path, m1_path, m2_path, avatar_path] if not p.exists()]
        if missing_assets:
            print(f"[ERR ] {folder.name}/{sku}: missing files: {', '.join(missing_assets)} — skipping this SKU")
            continue

        safe_product = "".join(product_name.split())
        out_name = f"{sku}_{safe_product}_24H.png"

        out_dir = Path("./output") / folder.name
        out_dir.mkdir(parents=True, exist_ok=True)

        out_path = out_dir / out_name

        try:
            canvas = bg_src.copy()

            # Open, scale to target heights, then rotate per angles
            avatar_img = rotate_image(scale_to_height(Image.open(avatar_path).convert("RGBA"), heights[IDX_AVATAR]), angles[IDX_AVATAR])
            m1_img     = rotate_image(scale_to_height(Image.open(m1_path).convert("RGBA"),      heights[IDX_M1]),     angles[IDX_M1])
            m2_img     = rotate_image(scale_to_height(Image.open(m2_path).convert("RGBA"),      heights[IDX_M2]),     angles[IDX_M2])
            a_img      = rotate_image(scale_to_height(Image.open(a_path).convert("RGBA"),       heights[IDX_A]),      angles[IDX_A])
            b_img      = rotate_image(scale_to_height(Image.open(b_path).convert("RGBA"),       heights[IDX_B]),      angles[IDX_B])

            # Paste in order using center coordinates (avatar → M1 → M2 → A → B)
            center_paste(canvas, avatar_img, centers[IDX_AVATAR])
            center_paste(canvas, m1_img,     centers[IDX_M1])
            center_paste(canvas, m2_img,     centers[IDX_M2])
            center_paste(canvas, b_img,      centers[IDX_B])
            center_paste(canvas, a_img,      centers[IDX_A])            

            # Foreground scissors last (already scaled & rotated)
            center_paste(canvas, scissors_img_final, centers[IDX_SCISSORS])

            canvas.save(out_path, format="PNG")
            print(f"[DONE] {folder.name}/{out_name}")

        except Exception as e:
            print(f"[ERR ] {folder.name}/{sku}: compose/save failed: {e}")

def main() -> int:
    ap = argparse.ArgumentParser(description="Generate 24H images from folder contents and Products.xlsx")
    ap.add_argument("--root", type=str, default=".", help="Root directory to scan for product folders (default: current directory)")
    args = ap.parse_args()

    root = Path(args.root).resolve()
    excel_path = root / EXCEL_NAME
    if not excel_path.exists():
        print(f"[FATAL] {EXCEL_NAME} not found beside script at: {excel_path}")
        return 2

    try:
        df = read_products(excel_path)
    except Exception as e:
        print(f"[FATAL] {e}")
        return 2

    grouped: Dict[str, pd.DataFrame] = dict(tuple(df.groupby(df["Folder"].astype(str).str.strip())))
    subdirs = [p for p in root.iterdir() if p.is_dir()]
    if not subdirs:
        print(f"[INFO] No subdirectories under {root}")
        return 0

    for folder in sorted(subdirs, key=lambda p: p.name.lower()):
        rows = grouped.get(folder.name, pd.DataFrame(columns=df.columns))
        process_folder(folder, rows)

    return 0

if __name__ == "__main__":
    raise SystemExit(main())
