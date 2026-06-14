"""Build the TMG Time Tracker app icon: the TMG mark + a clock.
Generates two on-brand variants as previews; the chosen one becomes icon-180/512.png.
Reuses the official TMG logo PNGs so the letterforms are authentic.
"""
import math, os
from PIL import Image, ImageDraw

PNG = r"C:\Users\nfpma\OneDrive - The Metals Group LLC\Documents\1_TMG Branding\LOGOs\Vector Outputs\PNG"
full = Image.open(PNG + r"\TMG_Logo_Master_FullColor_2048w.png").convert("RGBA")
white = Image.open(PNG + r"\TMG_Logo_SingleColor_White_2048w.png").convert("RGBA")

NAVY = (0, 49, 87, 255)     # #003157 — TMG brand navy
WHITE = (255, 255, 255, 255)

# --- isolate the "TMG" block (drop the "THE METALS GROUP LLC" tagline) ---
# The tagline sits below the biggest vertical gap in the artwork.
W, H = white.size
alpha = list(white.split()[3].getdata())
row_has = [max(alpha[y * W:(y + 1) * W]) > 20 for y in range(H)]
ys = [y for y in range(H) if row_has[y]]
top, max_gap, split = ys[0], 0, None
for a, b in zip(ys, ys[1:]):
    if b - a > max_gap:
        max_gap, split = b - a, a
bottom = (split + 1) if (split and max_gap > int(H * 0.02)) else ys[-1] + 1
strip = (0, top, W, bottom)
tmg_white = white.crop(strip)
bb = tmg_white.getbbox()
tmg_white = tmg_white.crop(bb)
tmg_full = full.crop(strip).crop(bb)
print("TMG crop strip", strip, "tight", bb, "->", tmg_white.size)


def draw_clock(d, cx, cy, r, color, hub_inner):
    d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=color, width=int(r * 0.11))
    for i in range(12):
        ang = math.radians(i * 30 - 90)
        outer = r - int(r * 0.05)
        inner = outer - (int(r * 0.17) if i % 3 == 0 else int(r * 0.10))
        w = int(r * 0.06) if i % 3 == 0 else int(r * 0.035)
        d.line([int(cx + math.cos(ang) * inner), int(cy + math.sin(ang) * inner),
                int(cx + math.cos(ang) * outer), int(cy + math.sin(ang) * outer)], fill=color, width=w)

    def hand(frac, deg, w):
        ang = math.radians(deg - 90)
        x, y = int(cx + math.cos(ang) * r * frac), int(cy + math.sin(ang) * r * frac)
        d.line([cx, cy, x, y], fill=color, width=w)
        cap = w // 2
        d.ellipse([x - cap, y - cap, x + cap, y + cap], fill=color)

    hand(0.74, 60, int(r * 0.05))    # minute hand -> 10:10 (classic, balanced)
    hand(0.52, 305, int(r * 0.075))  # hour hand
    hub = int(r * 0.075)
    d.ellipse([cx - hub, cy - hub, cx + hub, cy + hub], fill=color)
    h2 = int(hub * 0.45)
    d.ellipse([cx - h2, cy - h2, cx + h2, cy + h2], fill=hub_inner)


def make(bg, tmg, clock_color, hub_inner):
    S = 2048
    canvas = Image.new("RGBA", (S, S), bg)
    tw = int(S * 0.62)
    th = int(tw * tmg.height / tmg.width)
    canvas.alpha_composite(tmg.resize((tw, th), Image.LANCZOS), ((S - tw) // 2, int(S * 0.16)))
    draw_clock(ImageDraw.Draw(canvas), S // 2, int(S * 0.70), int(S * 0.20), clock_color, hub_inner)
    return canvas


variants = {
    "navy": make(NAVY, tmg_white, WHITE, NAVY),    # navy tile, white TMG, white clock
    "white": make(WHITE, tmg_full, NAVY, WHITE),   # white tile, full-color TMG (silver M), navy clock
}
out = os.path.join(os.path.dirname(__file__), "preview")
os.makedirs(out, exist_ok=True)
for name, img in variants.items():
    for sz in (512, 180):
        img.resize((sz, sz), Image.LANCZOS).convert("RGB").save(os.path.join(out, f"icon_{name}_{sz}.png"))
print("wrote previews to", out)
