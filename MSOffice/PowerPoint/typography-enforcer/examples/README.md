# Example — Global Typography Enforcer

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates both decks. |
| `sample_frankendeck.pptx` | A two-slide deck with deliberately mixed fonts (the "frankendeck") — the *deck you restyle*. |
| `sample_brand_template.pptx` | A one-slide reference deck with distinct title/body placeholder fonts — the *template you read styling from* (for the new template modes). |

### What's seeded, and why

The enforcer rewrites fonts **per text run**, in one of two modes, and always
preserves symbol/icon fonts. The sample is built to exercise every branch:

| Where | Text | Font | Role in the test |
|---|---|---|---|
| Slide 1 | `Quarterly Review` | **Calibri** | the font you'll target in the "specific" swap — should change |
| Slide 1 | `Prepared by the planning team` | **Arial** | a different font — must be left alone in "specific" mode |
| Slide 2 textbox | `This is Calibri. ` | **Calibri** | should change |
| Slide 2 textbox | `This is Arial. ` | **Arial** | left alone in "specific" mode |
| Slide 2 textbox | `This is Times. ` | **Times New Roman** | left alone in "specific" mode |
| Slide 2 textbox | `abc` | **Wingdings** | symbol font on the hard-coded blocklist — preserved in **both** modes |
| Slide 2 table (2×2) | four cells | Calibri / Arial / Calibri / Times | exercises the table-cell traversal |

The Wingdings run is the important one: those three letters render as
pictographs. Reflowing them to a text font would turn the icons into the
letters "abc", so the tool must never touch them — that's what the blocklist
guarantees.

## How to try it

1. Open `sample_frankendeck.pptx` in PowerPoint.
2. Run **Global Typography Enforcer** from ScripTree.

### Run A — "specific" swap (Calibri → Aptos)

| Field | Value |
|---|---|
| Replacement mode | `specific` |
| Target font to remove | `Calibri` |
| New font to apply | `Aptos` |
| Enforce on Slide Masters & Layouts | *(checked)* |
| Output format | `markdown` |

**Expected** — a sibling `sample_frankendeck_Restyled.pptx` in which:

* every **Calibri** run is now **Aptos** (title on slide 1, `This is Calibri.`
  run, both Calibri table cells);
* **Arial** and **Times New Roman** runs are **unchanged** — "specific" mode
  only touches the named source font;
* the **Wingdings** `abc` run is **unchanged** (preserved by the blocklist);
* the original deck is untouched.

### Run B — "all" mode (everything → Aptos)

| Field | Value |
|---|---|
| Replacement mode | `all` |
| Target font to remove | *(ignored / leave blank)* |
| New font to apply | `Aptos` |

**Expected** — in the `_Restyled` copy, **all text fonts** (Calibri, Arial,
*and* Times New Roman) become **Aptos** — but the **Wingdings** run is **still
preserved**. The blocklist wins even in "all" mode; that is the one invariant
that holds regardless of mode.

### Run C — read styles from a template (the `sample_brand_template.pptx` deck)

These two runs use the **Reference template deck** field. Open
`sample_frankendeck.pptx` as the deck to restyle, and point the template field
at `sample_brand_template.pptx`.

| Field | Value |
|---|---|
| Reference template deck (optional) | *(path to)* `sample_brand_template.pptx` |
| How to use the template | *(see the two sub-runs below)* |

When the template field is set, the manual **Replacement mode / source / target
font** fields are ignored.

* **C1 — "Apply the full template theme"** *(`theme`)* — the tool runs
  `work.ApplyTemplate(sample_brand_template.pptx)` on the working copy. The
  copy adopts the template deck's WHOLE theme (fonts + colours + masters). This
  is the mode the bundled sample exercises most faithfully (see the honesty
  note below) — `ApplyTemplate` copies the theme regardless of how it was
  authored.
* **C2 — "Use the template's theme fonts (headings + body)"** *(`fonts`)* —
  the tool reads the template's **theme** heading (major) + body (minor) fonts
  and sweeps the working copy so titles/subtitles get the heading font and all
  other text gets the body font (symbol fonts still preserved).

> **Honesty note — what the bundled sample does and does NOT exercise.**
> The `fonts` mode reads the template's **theme font scheme**
> (`SlideMaster.Theme.ThemeFontScheme.MajorFont/MinorFont`). `python-pptx` (used
> to build `sample_brand_template.pptx`) has **no API to author a real theme
> major/minor font scheme** — it can only set *direct* run fonts on the
> placeholders (which is what the generator does, so the deck plainly *looks*
> different: Georgia title, Verdana body). A deck built this way typically still
> carries the **default Office theme fonts** (Calibri Light / Calibri) in its
> theme. So on this sample, `fonts` mode may read back the Office defaults and
> demonstrate the plumbing rather than a dramatic change. On a **real corporate
> `.potx`** (which carries a proper theme font scheme) `fonts` mode reads the
> brand heading/body fonts as intended. If `fonts` mode reports the theme fonts
> could not be read, the tool returns `STATUS=NOTHEME` and suggests the
> full-theme option. **`theme` mode (ApplyTemplate) is the one this sample
> demonstrates faithfully.**

## What this demonstrates

* Per-**run** font rewriting (not per-shape or per-slide), so mixed-font
  paragraphs are handled correctly.
* The two modes: targeted single-font swap vs. blanket replace.
* Symbol/icon-font preservation via the hard-coded blocklist — the reason the
  Wingdings run survives both modes.
* Table-cell traversal (`Cell(r,c).Shape.TextFrame`).
* The work-on-a-copy safety model: output lands on `<name>_Restyled.pptx`,
  original untouched.

> Like every app in this catalog, the enforcer is pending live verification
> against a real PowerPoint. The scenario above describes the exact behaviour
> to expect from a deck with mixed fonts and a symbol-font run.
