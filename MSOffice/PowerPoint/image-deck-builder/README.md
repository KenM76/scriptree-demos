# Image-to-Slide Deck Builder (PowerPoint)

Build a **brand-new** PowerPoint deck from a folder of images — one image per
slide, each **scaled to fit the slide and centered**. Point it at a folder of
photos, screenshots, or scanned pages and it assembles a ready-to-review
slideshow in seconds.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

This is a **create-only** tool. Unlike the other PowerPoint apps in the catalog
(the Purger and the Typography Enforcer, which work on a copy of an *existing*
deck), this app **never opens, reads, or modifies any existing presentation** —
it makes a new one. There is therefore **no work-on-a-copy guard**; the only
collision risk is the output file name, which the shim de-duplicates.

---

## 1. What the user sees (end-user guide)

### The form fields

| Field | Type / widget | Default | Meaning |
|---|---|---|---|
| **Image folder** | path / `folder` (required) | — | The folder of images. Scanned **NON-recursively** (only this folder, not sub-folders) for files ending in `.jpg .jpeg .png .gif .bmp .tif .tiff` (case-insensitive). Images are added in **filename order**, one per slide. |
| **Output file** | string / `text` (optional) | *(blank)* | Where to save the new deck. **Blank** → `Slideshow.pptx` inside the image folder. A **bare name** (no folder) → saved inside the image folder. A **full path** → saved there. The name is forced to a `.pptx` extension. If a file of that name exists it is **NOT overwritten** — a ` (2)`, ` (3)`… suffix is added. |
| **Add filename captions** | boolean / `checkbox` | OFF | Add each image's filename (without its extension) as a caption text box across the bottom of every slide. |
| **Output format** | enum / `dropdown` | `markdown` | `Markdown` (paste into a ticket) or `Plain text` (aligned list). The result-summary rendering only. |

### What it does to your files

**Nothing to anything that already exists.** It creates one new `.pptx` at the
output path (de-duped so it never clobbers an earlier deck) and **opens it in a
window** so you can review it immediately. The deck is left open and is **not**
saved-over again, so you decide whether to keep it.

### Prerequisites

* PowerPoint is **running** (the tool builds the deck into that instance). No
  particular deck needs to be open — any other open deck is left untouched.
* The image folder exists and contains at least one supported image.
* ScripTree's bundled `lib/combridge/combridge.exe` is present.

---

## 2. The logic (reconstruct-the-tool spec)

A **Strategy-A shim**: a ScripTree form runs `build_image_deck.py`; the shim
**enumerates the image folder itself**, bakes the resulting absolute-path list
(plus the output path and flags) into a C# Roslyn script rendered from
`build_image_deck.csx.template`, then runs it through combridge, which owns the
live COM connection to PowerPoint.

* **combridge `run-script` has no argv channel** — a `.csx` only sees the plugin
  globals (`pptApp` / `pptPres`) and environment. Form values *and the image
  list* are baked in by replacing `__TOKEN__` placeholders.
* **combridge swallows the script's `return` value** — a clean run exits 0. So
  the script prints a first-line **sentinel** and the shim translates it into the
  process exit code.

### The shim-enumerates-the-folder-and-bakes-the-list design

The directory scan lives in the **Python shim**, not the `.csx`. Rationale:

1. **One tested place for file-system logic.** Path hygiene (resolving the
   output name, de-duping it, rejecting a missing folder) happens in Python
   **before any COM call** — fast to fail, easy to unit-test, no PowerPoint
   needed to validate the folder logic.
2. **A flat, deterministic `.csx`.** The generated script contains a literal C#
   `string[]` of absolute paths in filename order. There is no
   `Directory.GetFiles` quirk to reason about on the COM side; the script just
   iterates a list. The offline render-check can see exactly which slides will
   be built.
3. **Clean precondition split.** The shim catches a **missing folder**
   (`BADFOLDER`); the `.csx` catches a folder that exists but holds **no
   supported images** (`NOIMAGES`). Both emit the same sentinel shape, so the
   output pane is consistent regardless of which side caught it.

The scan itself: `folder.iterdir()` (NON-recursive), keep files whose
lower-cased suffix is in `{.jpg .jpeg .png .gif .bmp .tif .tiff}`, sort by
lower-cased filename. Each match is resolved to an absolute path and
`csharp_literal`-escaped into the array initialiser.

### Output-path resolution + de-dupe (shim)

```
blank            -> <folder>/Slideshow.pptx
bare name "x"    -> <folder>/x.pptx        (extension forced to .pptx)
full/rel path    -> as given, .pptx forced
```

Then, if the chosen path exists, append `" (2)"`, `" (3)"`, … before the
extension until a free name is found. The original deck is never overwritten.

### Sentinel + status → exit-code contract

First stdout line (from the shim on `BADFOLDER`, otherwise from the `.csx`):

```
__PPTIMG__ STATUS=<code> [slides=N failed=F out=PATH]
```

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Deck built; report follows (`slides=N failed=F out=PATH`). `failed` counts images skipped because `AddPicture`/layout threw. |
| `NOIMAGES` | 2 | Folder existed but held no supported image files. |
| `BADFOLDER` | 2 | The image folder does not exist / is not a directory (caught by the shim before PowerPoint is touched). |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes. |

### What the `.csx` does

1. **Guard:** baked-in `imagePaths.Length == 0` → `NOIMAGES` (return 0; shim → exit 2).
2. **DisplayAlerts off** (`pptApp.DisplayAlerts = PpAlertLevel.ppAlertsNone`,
   restored in a `finally`) — guards the `SaveAs` against a modal overwrite
   prompt that would block the automation host invisibly.
3. **Create the deck WITH a window:** `var prs = pptApp.Presentations.Add(MsoTriState.msoTrue)`.
   `msoTrue` = create it with a visible window so the user sees the result.
4. **Read the canvas size in points:** `float sw = prs.PageSetup.SlideWidth`,
   `float sh = prs.PageSetup.SlideHeight`. All positioning is in points.
5. **For each image path** (1-based insert index `idx`), inside a try/catch so a
   single bad image is skipped (`failed++`) rather than aborting the build:
   * **Blank slide:** `var slide = prs.Slides.Add(idx, PpSlideLayout.ppLayoutBlank); idx++;`
     — blank layout so no placeholders fight the centering.
   * **Insert at native size:**
     `var pic = slide.Shapes.AddPicture(path, LinkToFile:msoFalse, SaveWithDocument:msoTrue, Left:0, Top:0);`
     — `LinkToFile=msoFalse` embeds the bytes (portable deck); `SaveWithDocument=msoTrue`
     keeps the embedded copy in the file. Width/Height are **omitted** so the
     picture comes in at its real pixel-derived point size and `pic.Width` /
     `pic.Height` report true dimensions for the math.
   * **Scale-to-fit + center** (see §3 below):
     ```csharp
     pic.LockAspectRatio = MsoTriState.msoTrue;
     float scale = Math.Min(sw / pic.Width, sh / pic.Height);
     pic.Width = pic.Width * scale;          // Height follows (aspect locked)
     pic.Left  = (sw - pic.Width)  / 2;
     pic.Top   = (sh - pic.Height) / 2;
     ```
   * **Optional caption** (if `addCaptions`), in its own try/catch so a caption
     failure never loses the slide:
     ```csharp
     var tb = slide.Shapes.AddTextbox(
         MsoTextOrientation.msoTextOrientationHorizontal, 0, sh - 40, sw, 36);
     tb.TextFrame.TextRange.Text = <filename without extension>;
     ```
6. **Save + leave open:**
   `prs.SaveAs(outputPath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);`
   then **no `Close`** — the deck stays open in its window for review.
7. **Report:** sentinel
   `__PPTIMG__ STATUS=OK slides={added} failed={failed} out={outputPath}` then a
   markdown/plain-text summary (images found, slides added, any skipped, captions
   on/off, output path).

---

## 3. The scale-to-fit math (why `Math.Min`)

Each image is dropped in at its native size, then resized so it is as large as
possible while still fitting **entirely** inside the slide, with aspect ratio
preserved, then centered.

* `sw / pic.Width` is the factor that would make the image exactly the slide's
  **width**; `sh / pic.Height` the factor for the slide's **height**.
* Taking **`Math.Min`** of the two picks the *limiting* dimension: scaling by
  that factor makes one dimension touch the slide edge and leaves the other
  **inside** the slide (it can't overflow, because the larger factor would have
  overflowed the limiting dimension). The non-limiting dimension is then
  letter-boxed (top/bottom) or pillar-boxed (left/right) by the centering.
* `LockAspectRatio = msoTrue` means setting `pic.Width` proportionally adjusts
  `pic.Height`, so **one** assignment resizes without distortion. Then
  `Left = (sw - pic.Width)/2` and `Top = (sh - pic.Height)/2` center the result.
* `scale` may be **> 1** (a small image is enlarged to fill the slide) or **< 1**
  (a large image is shrunk to fit) — both are the correct "fit to slide".

Because slides are wide (16:9 by default in modern PowerPoint), a **portrait**
image will be height-limited (tall, pillar-boxed); a **landscape** image will be
width-limited (wide, letter-boxed). The example deck mixes aspect ratios on
purpose so this is visible.

---

## 4. Files in this app

| File | Role |
|---|---|
| `image-deck-builder.scriptree` | The form (4 params; a stacked-photos-becoming-a-slide PNG icon embedded — the "folder of images → slideshow" motif). |
| `image-deck-builder.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config (popups on, command-preview/extras off). |
| `build_image_deck.py` | The Strategy-A shim — **enumerates the folder**, resolves+dedupes the output path, bakes the image list, owns the exit code. |
| `build_image_deck.csx.template` | The Roslyn template with `__TOKEN__` placeholders (`__IMAGE_PATHS__` is the baked array body). |
| `README.md` | This file. |
| `examples/` | A generator + four labeled sample images + an example README. |

### argv contract (shim ⇆ form)

```
build_image_deck.py
  --source-folder  <folder>        (required; the path/folder picker, no_split)
  --output-file    <name-or-path>  (optional; blank => Slideshow.pptx in folder)
  [--add-captions]                 (flag; emitted only when ticked)
  --output-format  markdown|text
```

The two path/string values pass through `["--flag","{id}"]` token groups (the
`output_file` group **drops out entirely** when blank, so argparse uses its `""`
default). `add_captions` uses the conditional `{id?--flag}` token form.
`source_folder` and `output_file` are marked `no_split` so a value containing
spaces stays a single argv token.

---

## 5. PowerPoint COM facts this relies on

| Fact | Detail |
|---|---|
| **`Presentations.Add(MsoTriState.msoTrue)`** | Creates a new, empty presentation. `msoTrue` = **with a window** (visible) so the user sees the result; `msoFalse` would be headless. |
| **`Slides.Add(index, PpSlideLayout.ppLayoutBlank)`** | Inserts a blank slide at the 1-based `index`. Blank layout = no title/body placeholders. |
| **`Shapes.AddPicture(File, LinkToFile, SaveWithDocument, Left, Top[, Width, Height])`** | `LinkToFile=msoFalse` + `SaveWithDocument=msoTrue` **embeds** the image. Omitting Width/Height inserts at native size; the returned `Shape` exposes `.Width`/`.Height` in points for the scale math. |
| **`Shape.LockAspectRatio = msoTrue`** | Setting `.Width` then proportionally adjusts `.Height` — one assignment, no distortion. |
| **`Shapes.AddTextbox(MsoTextOrientation, Left, Top, Width, Height)`** | Creates a text box; `MsoTextOrientation.msoTextOrientationHorizontal` is the normal left-to-right orientation. Text set via `.TextFrame.TextRange.Text`. |
| **`Presentation.SaveAs(path, PpSaveAsFileType.ppSaveAsOpenXMLPresentation)`** | Writes a modern `.pptx`. We do **not** `Close` afterwards — the deck stays open for review. |
| **`MsoTriState` / `MsoTextOrientation`** | From `Microsoft.Office.Core`; the `Pp*` enums (`PpSlideLayout`, `PpSaveAsFileType`, `PpAlertLevel`) from `Microsoft.Office.Interop.PowerPoint`. **Both are in the PowerPoint plugin's default ScriptUsings** — no extra `using` needed. |
| **DisplayAlerts** | `pptApp.DisplayAlerts = PpAlertLevel.ppAlertsNone` around the `SaveAs` to suppress a modal overwrite prompt (which would block the automation host invisibly). Restored in `finally`. |

---

## 6. Editing / maintenance notes

* **Stays create-only.** The value of this tool is that it can never harm an
  existing deck. Do not add an "append to the open deck" mode here — that
  reintroduces the work-on-a-copy question; make it a separate tool instead.
* **Enumerate in the shim, not the `.csx`.** Keep the folder scan, extension
  filter, sort, and output-path dedupe in Python so they're testable offline and
  the generated script stays a flat, auditable list.
* **Per-image try/catch is load-bearing.** One corrupt/locked image must not
  abort the whole build; it's counted as `failed` and the run is still `OK`
  (exit 0) as long as the deck saved.
* **Captions are cosmetic** — their own try/catch ensures a caption failure
  never costs you the slide.
* **`failed == imagePaths.Length`?** Every image was unreadable; the deck still
  saves (empty) and reports `slides=0 failed=N`. This is an honest result, not a
  precondition error, so it stays exit 0. (Adjust here if you'd rather treat an
  all-failed build as exit 2.)
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree validate <path>`.
* **Offline render-check:** render the template with a sample 2–3 path list
  (include an awkward path with a quote/backslash) and grep for `__[A-Z_]+__` —
  only the `__PPTIMG__` sentinel should match; anything else is an unfilled
  placeholder.
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path.
