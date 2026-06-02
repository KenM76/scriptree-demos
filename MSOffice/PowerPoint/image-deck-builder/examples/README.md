# Example — Image-to-Slide Deck Builder

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates the four sample images. |
| `sample_images/` | Four labeled PNGs with **different aspect ratios**, created by the generator. |

### What's generated, and why

The builder scans a folder for images and makes **one scaled-to-fit, centered
slide per image** in a new deck. The four samples deliberately span the aspect
ratios so the fit-and-center math is visible in the result:

| File | Size | Shape | What you should see on its slide |
|---|---|---|---|
| `Slide 1.png` | 1600×900 | landscape 16:9 | Fills the slide nearly edge-to-edge (it matches the slide's own 16:9 shape). |
| `Slide 2.png` | 900×1600 | portrait 9:16 | Tall and centered, **pillar-boxed** (empty bars left and right). |
| `Slide 3.png` | 1200×1200 | square 1:1 | Centered square, boxed on the wider (horizontal) axis. |
| `Slide 4.png` | 640×360 | small landscape | **Upscaled** to fill the slide width (scale > 1) — proving small images are enlarged to fit, not left tiny. |

Each image is a solid color with a big white `Slide N` label and its pixel
dimensions, so in the produced deck you can tell at a glance which image landed
where and confirm none are stretched or cropped.

## How to try it

1. Generate the images: `python make_example.py` (creates `sample_images/`).
2. Make sure **PowerPoint is running** (any deck, or none — it won't be touched).
3. Run **Image-to-Slide Deck Builder** from ScripTree.

### Run A — default (no captions)

| Field | Value |
|---|---|
| Image folder | `…/examples/sample_images` |
| Output file | *(blank)* |
| Add filename captions | *(unchecked)* |
| Output format | `markdown` |

**Expected** — a new `Slideshow.pptx` is created **inside `sample_images/`**,
opened in a PowerPoint window, with **4 slides** in filename order
(`Slide 1` → `Slide 4`), each image **centered and scaled to fit**:

* slide 1 nearly fills the canvas;
* slide 2 is tall and pillar-boxed;
* slide 3 is a centered square;
* slide 4 is the small image **enlarged** to fit.

The summary reports `Images found: 4`, `Slides added: 4`, `Captions: (none)`,
and the saved path. Nothing else is modified.

### Run B — with captions

Re-run with **Add filename captions** = *checked*.

**Expected** — same 4 slides, but each now has a caption text box across the
bottom showing the filename without its extension (`Slide 1`, `Slide 2`, …).
Because a `Slideshow.pptx` already exists from Run A, the new deck is saved as
**`Slideshow (2).pptx`** — the de-dupe in action; your Run-A deck is preserved.

## What this demonstrates

* Building a brand-new deck from a folder of images, one image per slide.
* The **scale-to-fit + center** math across mixed aspect ratios (landscape,
  portrait, square) and the **upscaling** of an undersized image.
* The optional filename captions.
* The **create-only** safety model and the **output-name de-dupe** (`Slideshow
  (2).pptx`), so re-runs never overwrite an earlier deck.

> Like every app in this catalog, the builder is pending live verification
> against a real PowerPoint. The `AddPicture` / `AddTextbox` argument shapes and
> the `Presentations.Add(msoTrue)` with-window create are the items flagged for
> that live check.
