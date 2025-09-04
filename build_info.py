#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build info image pipeline (v3) — Python 3.12.3

Fixes & Features
- FIX: previous syntax error removed; file rewritten cleanly.
- Composites overlays onto background per mapping:
  * {SKU}_M1.png -> positions M1 and M3
  * {SKU}_M2.png -> position M2
- Output path: ./output/<Folder>/{SKU}_{ProductNameNoSpaces}_info.png
- Reads constraints/config:
  * positions_info.txt : 3 lines "x,y" for M1/M2/M3 centers
  * rotate_info.txt    : 3 lines angle(deg); negative=CCW, positive=CW
  * maxheight_info.txt : 3 lines (optional) max heights for M1/M2/M3
- Case-tolerant lookup for overlay filenames (M1/m1, M2/m2).

Dependencies: Pillow, pandas, openpyxl
    pip install pillow pandas openpyxl
"""
from __future__ import annotations

import sys
from pathlib import Path
import pandas as pd
from PIL import Image

SCRIPT_DIR = Path.cwd()
EXCEL_PATH = SCRIPT_DIR / "Products.xlsx"
OUTPUT_ROOT = SCRIPT_DIR / "output"

# -------------------- Logging helpers --------------------

def info(msg: str) -> None:
    print(f"[INFO] {msg}")


def warn(msg: str) -> None:
    print(f"[WARN] {msg}")


def err(msg: str) -> None:
    print(f"[ERROR] {msg}")

# -------------------- Config readers --------------------

def read_positions_file(p: Path) -> dict[str, tuple[float, float]] | None:
    keys = ["M1", "M2", "M3"]
    try:
        lines = p.read_text(encoding="utf-8").strip().splitlines()
    except FileNotFoundError:
        warn(f"Missing positions file: {p}")
        return None
    except Exception as e:
        err(f"Failed reading positions file {p}: {e}")
        return None
    if len(lines) < 3:
        warn(f"positions_info.txt has {len(lines)} lines; expected 3. Path={p}")
        return None
    out: dict[str, tuple[float, float]] = {}
    for i, k in enumerate(keys):
        raw = lines[i].strip().replace("﻿", "")
        try:
            xs, ys = [t.strip() for t in raw.split(",", 1)]
            out[k] = (float(xs), float(ys))
        except Exception:
            warn(f"Bad position line {i+1} in {p}: '{raw}' (expected 'x,y')")
            return None
    return out


def read_rotate_file(p: Path) -> dict[str, float] | None:
    keys = ["M1", "M2", "M3"]
    try:
        lines = p.read_text(encoding="utf-8").strip().splitlines()
    except FileNotFoundError:
        warn(f"Missing rotate file: {p}")
        return None
    except Exception as e:
        err(f"Failed reading rotate file {p}: {e}")
        return None
    if len(lines) < 3:
        warn(f"rotate_info.txt has {len(lines)} lines; expected 3. Path={p}")
        return None
    out: dict[str, float] = {}
    for i, k in enumerate(keys):
        raw = lines[i].strip().replace("﻿", "")
        try:
            out[k] = float(raw)
        except Exception:
            warn(f"Bad rotate line {i+1} in {p}: '{raw}' (expected a number)")
            return None
    return out


def read_maxheight_file(p: Path) -> dict[str, float] | None:
    keys = ["M1", "M2", "M3"]
    if not p.exists():
        return None
    try:
        lines = p.read_text(encoding="utf-8").strip().splitlines()
    except Exception as e:
        warn(f"Failed reading maxheight file {p}: {e}")
        return None
    if len(lines) < 3:
        warn(f"maxheight_info.txt has {len(lines)} lines; expected 3. Path={p}")
        return None
    out: dict[str, float] = {}
    for i, k in enumerate(keys):
        raw = lines[i].strip().replace("﻿", "")
        try:
            h = float(raw)
            if h <= 0:
                raise ValueError
            out[k] = h
        except Exception:
            warn(f"Bad maxheight line {i+1} in {p}: '{raw}' (expected positive number)")
            return None
    return out

# -------------------- Excel loader --------------------

def load_products(excel_path: Path) -> pd.DataFrame:
    if not excel_path.exists():
        err(f"Products.xlsx not found at: {excel_path}")
        sys.exit(1)
    try:
        df = pd.read_excel(excel_path, engine="openpyxl")
    except Exception as e:
        err(f"Failed to read Excel '{excel_path}': {e}")
        sys.exit(1)
    required = {"Folder", "SKU", "ProductName"}
    missing = required.difference(df.columns)
    if missing:
        err(f"Products.xlsx missing columns: {sorted(missing)}. Present: {list(df.columns)}")
        sys.exit(1)
    return df.dropna(subset=list(required))

# -------------------- Imaging helpers --------------------

def find_overlay_path(folder_path: Path, sku: str, variant: str) -> Path | None:
    """Case-tolerant search for {SKU}_{variant}.png (e.g., M1/m1)."""
    candidates = [
        folder_path / f"{sku}_{variant}.png",
        folder_path / f"{sku}_{variant.lower()}.png",
        folder_path / f"{sku}_{variant.upper()}.png",
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


def paste_center_rotated(
    bg: Image.Image,
    fg_path: Path,
    center: tuple[float, float],
    angle_deg: float,
    max_val: float | None = None,
) -> None:
    """Paste fg onto bg centered at `center` after optional proportional scaling and rotation.
    Scaling rule: if original fg.height > fg.width -> constrain height to `max_val` (if provided);
    else constrain width to `max_val`. Scaling occurs BEFORE rotation.
    Spec: positive angle = clockwise; Pillow positive rotates CCW -> use -angle.
    """
    if not fg_path.exists():
        raise FileNotFoundError(f"Missing overlay image: {fg_path}")
    with Image.open(fg_path).convert("RGBA") as fg:
        if max_val is not None and max_val > 0:
            w0, h0 = fg.size
            if h0 > w0:
                # portrait: limit height
                limit = max_val
                if h0 > limit:
                    scale = limit / float(h0)
                else:
                    scale = 1.0
            else:
                # landscape/square: limit width
                limit = max_val
                if w0 > limit:
                    scale = limit / float(w0)
                else:
                    scale = 1.0
            if scale < 1.0:
                new_w = max(1, int(round(w0 * scale)))
                new_h = max(1, int(round(h0 * scale)))
                fg = fg.resize((new_w, new_h), resample=Image.LANCZOS)
        rotated = fg.rotate(-angle_deg, expand=True, resample=Image.BICUBIC)
        rw, rh = rotated.size
        cx, cy = center
        left = int(round(cx - rw / 2))
        top = int(round(cy - rh / 2))
        bg.paste(rotated, (left, top), rotated)

# -------------------- Row processing --------------------

def process_row(base_dir: Path, folder: str, sku: str, product_name: str) -> None:
    folder_path = base_dir / folder
    if not folder_path.exists() or not folder_path.is_dir():
        warn(f"Folder not found or not a directory: {folder_path}")
        return

    bg_path = folder_path / "background_info.png"
    if not bg_path.exists():
        warn(f"Missing background_info.png, skip: {folder_path}")
        return

    positions = read_positions_file(folder_path / "positions_info.txt")
    rotations = read_rotate_file(folder_path / "rotate_info.txt")
    maxheights = read_maxheight_file(folder_path / "maxheight_info.txt")
    if positions is None or rotations is None:
        warn(f"Invalid positions/rotations; skip: {folder_path}")
        return

    m1_path = find_overlay_path(folder_path, sku, "M1")
    m2_path = find_overlay_path(folder_path, sku, "M2")
    missing = []
    if m1_path is None:
        missing.append(f"{sku}_M1.png")
    if m2_path is None:
        missing.append(f"{sku}_M2.png")
    if missing:
        warn(f"Missing required image(s) in {folder_path}: {', '.join(missing)}. Skipping.")
        return

    safe_product = "".join(str(product_name).split())
    out_dir = OUTPUT_ROOT / folder
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{sku}_{safe_product}_info.png"

    try:
        with Image.open(bg_path).convert("RGBA") as bg:
            # Max heights per placement (optional)
            mh1 = maxheights.get("M1") if isinstance(maxheights, dict) else None
            mh2 = maxheights.get("M2") if isinstance(maxheights, dict) else None
            mh3 = maxheights.get("M3") if isinstance(maxheights, dict) else None

            # Layering order (bottom -> middle -> top): M3 -> M2 -> M1
            # Bottom: place M3 using M1 image
            try:
                paste_center_rotated(bg, m1_path, positions["M3"], rotations["M3"], mh3)
            except KeyError:
                warn(f"Missing M3 center/angle; skip M1@M3 (bottom) in {folder_path}")

            # Middle: M2
            try:
                paste_center_rotated(bg, m2_path, positions["M2"], rotations["M2"], mh2)
            except KeyError:
                warn(f"Missing M2 center/angle; skip M2@M2 (middle) in {folder_path}")

            # Top: M1
            try:
                paste_center_rotated(bg, m1_path, positions["M1"], rotations["M1"], mh1)
            except KeyError:
                warn(f"Missing M1 center/angle; skip M1@M1 (top) in {folder_path}")

            bg.save(out_path, format="PNG")
            info(f"Created: {out_path.relative_to(SCRIPT_DIR)}")
    except Exception as e:
        err(f"Failed composing for folder '{folder}' (SKU={sku}): {e}")
        return

# -------------------- Main --------------------

def main() -> None:
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    df = load_products(EXCEL_PATH)

    processed = 0
    for idx, row in df.iterrows():
        folder = str(row["Folder"]).strip()
        sku = str(row["SKU"]).strip()
        pname = str(row["ProductName"]).strip()
        if not folder or not sku:
            warn(f"Row {idx}: empty Folder/SKU, skipping")
            continue
        info(f"Processing Folder='{folder}', SKU='{sku}' ...")
        process_row(SCRIPT_DIR, folder, sku, pname)
        processed += 1

    info(f"Done. Rows visited: {processed}")


if __name__ == "__main__":
    main()
