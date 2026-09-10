# -*- coding: utf-8 -*-
"""Keep source KV uncropped; dissolve it into white so the extend has no seam."""
from pathlib import Path

import numpy as np
from PIL import Image, ImageFilter

ROOT = Path(__file__).resolve().parent
TARGET = (1536, 1386)
JOBS = {
    "merchant": {
        "src": ROOT / "kv-previews" / "kv-merchant-user-source.png",
        "base": ROOT / "kv-merchant.png",
        "extend": ROOT / "kv-merchant-extend.png",
    },
    "partner": {
        "src": ROOT / "kv-previews" / "kv-partner-user-source.png",
        "base": ROOT / "kv-partner.png",
        "extend": ROOT / "kv-partner-extend.png",
    },
}


def smoothstep(x: np.ndarray) -> np.ndarray:
    x = np.clip(x, 0.0, 1.0)
    return x * x * (3.0 - 2.0 * x)


def ramp(height: int, start: float, end: float) -> np.ndarray:
    ys = np.arange(height, dtype=np.float32)
    return smoothstep((ys - start) / max(end - start, 1.0))[:, None, None]


def extend_to_size(src: Path, target: tuple[int, int], base_dst: Path, extend_dst: Path) -> None:
    target_w, target_h = target
    im = Image.open(src).convert("RGB")
    base_h = round(im.size[1] * target_w / im.size[0])
    im = im.resize((target_w, base_h), Image.Resampling.LANCZOS)
    im.save(base_dst)

    w, h = im.size
    extra = max(target_h - h, 8)
    total_h = h + extra
    canvas = Image.new("RGB", (w, total_h), (255, 255, 255))
    canvas.paste(im, (0, 0))

    blur = canvas.filter(ImageFilter.GaussianBlur(radius=56))
    haze = canvas.filter(ImageFilter.GaussianBlur(radius=90))
    sharp = np.asarray(canvas, dtype=np.float32)
    soft = np.asarray(blur, dtype=np.float32)
    mist = np.asarray(haze, dtype=np.float32)

    # Soften well above the original bottom so the join never reads as a line.
    blur_m = ramp(total_h, h * 0.38, h * 0.88)
    mixed = sharp * (1.0 - blur_m) + soft * blur_m
    haze_m = ramp(total_h, h * 0.52, h + extra * 0.35)
    mixed = mixed * (1.0 - haze_m) + mist * haze_m

    # White veil starts mid-photo and is nearly opaque by the original bottom.
    veil_m = ramp(total_h, h * 0.50, h + extra * 0.42)
    white = np.array([255.0, 255.0, 255.0], dtype=np.float32)
    out = mixed * (1.0 - veil_m) + white * veil_m

    Image.fromarray(np.clip(out, 0, 255).astype(np.uint8)).save(extend_dst)
    print(f"base {im.size} -> extend {out.shape[1], out.shape[0]}")


if __name__ == "__main__":
    import sys

    names = sys.argv[1:] or ["merchant"]
    for name in names:
        job = JOBS[name]
        extend_to_size(job["src"], TARGET, job["base"], job["extend"])
