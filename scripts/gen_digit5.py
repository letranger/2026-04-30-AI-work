#!/usr/bin/env python3
"""Generate two demo images for the AI=function explanation:

1. img/04-digit5-pixels.png  — handwritten "5" rendered side-by-side
   with its 28x28 pixel-value matrix (left: image, right: numbers).
2. img/04-digit5-prob.png    — f(image) → probability bar chart for
   classes 0-9 with class 5 winning at 0.72.

Run:
    uv run --with numpy --with matplotlib scripts/gen_digit5.py
"""

import os
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import font_manager

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "img")
os.makedirs(OUT_DIR, exist_ok=True)


# ── 1. Hand-craft a 28x28 "5" ─────────────────────────────────
# Design as a 28x28 binary mask, then smooth with a small box blur
# to produce gradient pixel values (0–255).

def craft_digit_5():
    img = np.zeros((28, 28), dtype=np.float32)

    # Top horizontal stroke (rows 4–7, cols 6–20)
    img[4:7, 6:21] = 1.0
    # Left vertical stroke down to middle (rows 6–14, cols 6–9)
    img[6:14, 6:9] = 1.0
    # Middle horizontal-ish bend (rows 13–16, cols 6–18)
    img[13:16, 6:18] = 1.0
    # Right side curve down (rows 14–22, cols 16–20)
    img[14:22, 17:21] = 1.0
    # Lower diagonal/hook back to left (rows 21–24, cols 6–19)
    img[21:24, 6:20] = 1.0
    # Slight bottom-left anchor to suggest closure
    img[20:24, 6:9] = 1.0

    # Add slight handwritten irregularity: random small noise + edge softening
    rng = np.random.default_rng(seed=5)
    noise = rng.normal(0, 0.06, img.shape)
    img = np.clip(img + noise * (img > 0.3), 0, 1)

    # Box blur 3x3 twice → smoother gradient at edges
    def box_blur(a):
        padded = np.pad(a, 1, mode="constant")
        out = np.zeros_like(a)
        for dy in range(3):
            for dx in range(3):
                out += padded[dy:dy + a.shape[0], dx:dx + a.shape[1]]
        return out / 9.0

    img = box_blur(img)
    img = box_blur(img)

    # Boost contrast a bit (gamma < 1 brightens mid-tones)
    img = img ** 0.85

    # Scale to 0–255 ints
    img = np.clip(img * 255, 0, 255).astype(np.int32)
    return img


# ── 2. Plot image 1: pixels left + matrix right ────────────────
def plot_pixels_and_matrix(arr, out_path):
    fig, (ax_img, ax_mat) = plt.subplots(
        1, 2, figsize=(13, 6), gridspec_kw={"width_ratios": [1, 1.7]}
    )

    # Left: image (matplotlib imshow with axis ticks 0–25)
    ax_img.imshow(arr, cmap="gray", vmin=0, vmax=255)
    ax_img.set_xticks(range(0, 28, 5))
    ax_img.set_yticks(range(0, 28, 5))
    ax_img.tick_params(labelsize=11)
    for spine in ax_img.spines.values():
        spine.set_visible(False)

    # Right: numeric matrix
    ax_mat.set_xlim(-0.5, 28)
    ax_mat.set_ylim(28, -0.5)
    ax_mat.axis("off")

    # Use a monospaced font for the grid of numbers
    mono = "Menlo" if "Menlo" in [f.name for f in font_manager.fontManager.ttflist] else "monospace"

    # Build each row as a single string for nicer alignment
    for r in range(28):
        if r == 0:
            prefix = "[["
        else:
            prefix = " ["
        suffix = "]]" if r == 27 else "]"
        cells = " ".join(f"{v:3d}" for v in arr[r])
        ax_mat.text(0, r, f"{prefix}{cells}{suffix}", fontsize=6.5,
                    family=mono, va="center", color="black")

    fig.tight_layout()
    fig.savefig(out_path, dpi=180, bbox_inches="tight",
                facecolor="white")
    plt.close(fig)
    print(f"saved: {out_path}")


# ── 3. Plot image 2: f(digit) → probability bars ──────────────
def plot_probabilities(arr, out_path):
    probs = [0.01, 0.02, 0.03, 0.05, 0.04, 0.72, 0.04, 0.04, 0.03, 0.02]

    fig = plt.figure(figsize=(12, 6))
    # 5 columns: "f(" | digit | ")" | arrow | bars
    gs = fig.add_gridspec(
        1, 5, width_ratios=[0.35, 0.55, 0.25, 0.55, 1.6], wspace=0.0
    )

    # Col 0: "f("
    ax_f = fig.add_subplot(gs[0, 0])
    ax_f.axis("off")
    ax_f.text(0.5, 0.5, "f(", fontsize=72, va="center", ha="center",
              family="serif")

    # Col 1: digit image
    ax_img = fig.add_subplot(gs[0, 1])
    ax_img.imshow(arr, cmap="gray", vmin=0, vmax=255)
    ax_img.axis("off")

    # Col 2: ")"
    ax_p = fig.add_subplot(gs[0, 2])
    ax_p.axis("off")
    ax_p.text(0.5, 0.5, ")", fontsize=72, va="center", ha="center",
              family="serif")

    # Col 3: arrow
    ax_arrow = fig.add_subplot(gs[0, 3])
    ax_arrow.axis("off")
    ax_arrow.annotate("", xy=(0.95, 0.5), xytext=(0.05, 0.5),
                      arrowprops=dict(arrowstyle="->", lw=2.5, color="black"))

    # Col 4: horizontal probability bars
    ax_right = fig.add_subplot(gs[0, 4])
    ax_right.set_xlim(0, 1.0)
    ax_right.set_ylim(-0.5, 9.5)
    ax_right.invert_yaxis()
    ax_right.axis("off")

    for i, p in enumerate(probs):
        ax_right.text(-0.05, i, f"{i}:", fontsize=18, va="center", ha="right",
                      family="sans-serif")
        ax_right.barh(i, p, height=0.55, color="#2A8ACB", edgecolor="none")
        ax_right.text(p + 0.015, i, f"{p:.2f}", fontsize=16, va="center",
                      ha="left", family="sans-serif")

    fig.savefig(out_path, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"saved: {out_path}")


# ── main ─────────────────────────────────────────────────────
arr = craft_digit_5()
plot_pixels_and_matrix(arr, os.path.join(OUT_DIR, "04-digit5-pixels.png"))
plot_probabilities(arr, os.path.join(OUT_DIR, "04-digit5-prob.png"))
