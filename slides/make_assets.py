"""背景素材を配色ごとに生成する。

pptxgenjs はグラデーション塗りに対応しないため、背景に敷く画像として書き出す。
assets/<palette>/bg-{title,body,accent,closing}.png を作る。
"""

import os

import numpy as np
from PIL import Image, ImageFilter

W, H = 1920, 1080

# 各配色の「主色・副色・和らげ色」と、下地となる白の色味
PALETTES = {
    "pink": {
        "base": (255, 250, 251),
        "base_body": (255, 252, 252),
        "base_accent": (255, 249, 250),
        "veil": (255, 253, 253),
        "main": (244, 205, 213),
        "main_warm": (247, 219, 216),
        "sub": (214, 226, 213),
        "warm": (250, 240, 226),
    },
    "blue": {
        "base": (250, 252, 255),
        "base_body": (251, 253, 255),
        "base_accent": (249, 252, 255),
        "veil": (253, 254, 255),
        "main": (198, 220, 238),
        "main_warm": (208, 226, 243),
        "sub": (199, 223, 229),
        "warm": (224, 236, 247),
    },
}


def canvas(base):
    return np.ones((H, W, 3), dtype=np.float64) * np.array(base, dtype=np.float64)


def blob(arr, cx, cy, rx, ry, color, strength, softness=1.0):
    """楕円状のにじみを合成する。cx/cy/rx/ry は 0-1 の相対値。"""
    y, x = np.mgrid[0:H, 0:W]
    nx = (x - cx * W) / (rx * W)
    ny = (y - cy * H) / (ry * H)
    d = np.sqrt(nx**2 + ny**2)
    mask = np.clip(1.0 - d, 0.0, 1.0) ** (2.0 * softness)
    mask = mask[:, :, None] * strength
    col = np.array(color, dtype=np.float64)
    return arr * (1 - mask) + col * mask


def save(arr, out_dir, name, blur=130):
    img = Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8))
    img = img.filter(ImageFilter.GaussianBlur(blur))
    path = f"{out_dir}/{name}"
    img.save(path, quality=95)
    print(path)


def build(palette, p):
    out = f"assets/{palette}"
    os.makedirs(out, exist_ok=True)
    MAIN, WARM_MAIN, SUB, WARM, VEIL = p["main"], p["main_warm"], p["sub"], p["warm"], p["veil"]

    # タイトル用: 四隅に淡いにじみを置き、中央は文字が乗るよう明るく保つ
    a = canvas(p["base"])
    a = blob(a, 0.08, 0.12, 0.42, 0.55, MAIN, 0.85)
    a = blob(a, 0.92, 0.10, 0.38, 0.48, WARM_MAIN, 0.70)
    a = blob(a, 0.88, 0.92, 0.45, 0.55, MAIN, 0.75)
    a = blob(a, 0.14, 0.94, 0.40, 0.45, SUB, 0.42)
    a = blob(a, 0.62, 0.86, 0.30, 0.32, WARM, 0.45)
    a = blob(a, 0.50, 0.45, 0.72, 0.62, VEIL, 0.62, softness=2.4)
    save(a, out, "bg-title.png")

    # 本文用: ごく薄く、対角の角だけに気配を残す
    a = canvas(p["base_body"])
    a = blob(a, 0.02, 0.04, 0.30, 0.38, MAIN, 0.40)
    a = blob(a, 0.99, 0.97, 0.32, 0.40, WARM_MAIN, 0.34)
    a = blob(a, 0.55, 0.50, 0.80, 0.75, VEIL, 0.52, softness=2.4)
    save(a, out, "bg-body.png")

    # 節目用: 本文よりやや色を強め、章の切り替わりを示す
    a = canvas(p["base_accent"])
    a = blob(a, 0.10, 0.85, 0.45, 0.50, MAIN, 0.60)
    a = blob(a, 0.90, 0.15, 0.42, 0.48, SUB, 0.40)
    a = blob(a, 0.50, 0.45, 0.75, 0.68, VEIL, 0.55, softness=2.4)
    save(a, out, "bg-accent.png")

    # 締め用: 下辺から淡く立ち上がる
    a = canvas(p["base"])
    a = blob(a, 0.20, 1.02, 0.55, 0.55, MAIN, 0.80)
    a = blob(a, 0.78, 1.00, 0.50, 0.48, WARM_MAIN, 0.65)
    a = blob(a, 0.50, 0.96, 0.40, 0.30, WARM, 0.40)
    a = blob(a, 0.95, 0.05, 0.30, 0.35, SUB, 0.30)
    a = blob(a, 0.45, 0.35, 0.80, 0.62, VEIL, 0.58, softness=2.4)
    save(a, out, "bg-closing.png")


for name, spec in PALETTES.items():
    build(name, spec)
