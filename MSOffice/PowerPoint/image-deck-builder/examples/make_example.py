#!/usr/bin/env python3
"""Generate the sample images for the Image-to-Slide Deck Builder.

Run:  python make_example.py
Produces:  sample_images/Slide 1.png .. Slide 4.png  (next to this script)

WHAT THE SAMPLE CONTAINS (and why)
----------------------------------
Four distinct, clearly-labeled PNGs with DIFFERENT aspect ratios, so that the
deck builder's "scale to fit the slide and center" behaviour is visible:

* Slide 1 - 1600x900  landscape (16:9)  - fills the slide width edge to edge.
* Slide 2 -  900x1600 portrait (9:16)   - tall; pillar-boxed (bars left/right).
* Slide 3 - 1200x1200 square (1:1)      - centered, boxed on the wider axis.
* Slide 4 -  640x360  small landscape   - UPSCALED to fit (scale > 1).

Each image is a flat colored background with its big white label centered, so
in the produced deck you can instantly tell which image landed on which slide
and confirm none are stretched or cropped.

Point the Image-to-Slide Deck Builder at the generated ``sample_images/``
folder and you get a 4-slide ``Slideshow.pptx`` with each image centered and
scaled to fit. ASCII-only output here (the Bash tool encodes cp1252 and chokes
on non-ASCII in print()).
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

# (label, width, height, background RGB)
IMAGES = [
    ("Slide 1", 1600, 900, (37, 99, 235)),    # landscape 16:9 - blue
    ("Slide 2", 900, 1600, (22, 163, 74)),     # portrait 9:16 - green
    ("Slide 3", 1200, 1200, (217, 119, 6)),    # square 1:1 - amber
    ("Slide 4", 640, 360, (190, 24, 93)),      # small landscape - magenta
]


def load_font(size):
    """Best-effort large TrueType font; fall back to PIL's bitmap default."""
    for name in ("arial.ttf", "DejaVuSans.ttf", "LiberationSans-Regular.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def draw_label(img, text):
    d = ImageDraw.Draw(img)
    w, h = img.size
    font = load_font(max(48, min(w, h) // 6))
    # Measure with textbbox so the label is genuinely centered.
    bbox = d.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = (w - tw) / 2 - bbox[0]
    y = (h - th) / 2 - bbox[1]
    # A subtle dimension subtitle under the main label.
    d.text((x, y), text, fill=(255, 255, 255), font=font)
    sub = f"{w} x {h}"
    sfont = load_font(max(24, min(w, h) // 16))
    sbbox = d.textbbox((0, 0), sub, font=sfont)
    sw_, sh_ = sbbox[2] - sbbox[0], sbbox[3] - sbbox[1]
    d.text(((w - sw_) / 2 - sbbox[0], y + th + th // 3), sub,
           fill=(255, 255, 255), font=sfont)


def main():
    out_dir = Path(__file__).resolve().parent / "sample_images"
    out_dir.mkdir(exist_ok=True)
    for label, w, h, bg in IMAGES:
        img = Image.new("RGB", (w, h), bg)
        draw_label(img, label)
        path = out_dir / f"{label}.png"
        img.save(str(path), format="PNG")
        print(f"wrote {path}  ({w}x{h})")
    print(f"done: {len(IMAGES)} images in {out_dir}")


if __name__ == "__main__":
    main()
