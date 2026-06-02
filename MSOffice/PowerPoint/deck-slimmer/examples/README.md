# Example — Slide Deck Slimmer

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates the deck. |
| `sample_bloated.pptx` | A two-slide deck built on the default Office Theme template, which carries 11 custom layouts; only 2 are used, so 9 are unused dead weight. |

### What's seeded, and why

The Slimmer removes **unused custom slide layouts**. The default python-pptx
template ("Office Theme") ships with **11 built-in slide layouts**. The sample
adds just **two slides**, each using **one** layout:

| Slide | Layout used |
|---|---|
| 1 | `Title Slide` (layout index 0) |
| 2 | `Title and Content` (layout index 1) |

That leaves **9 unused layouts** in the deck (`Section Header`, `Two Content`,
`Comparison`, `Title Only`, `Blank`, `Content with Caption`,
`Picture with Caption`, `Title and Vertical Text`, `Vertical Title and Text`).
python-pptx keeps **all** of the template's layouts in the saved file (it does
not prune unused ones), and PowerPoint exposes them via
`Design.SlideMaster.CustomLayouts` — exactly the collection the Slimmer
iterates — so these unused layouts are visible to, and removable by, the tool.

Verified after generation (reopened with python-pptx):

* **1** slide master, **11** custom layouts, all with distinct names.
* **2** slides, using `Title Slide` and `Title and Content`.
* **9** layouts unused → expected to be removed.

## How to try it

1. Open `sample_bloated.pptx` in PowerPoint.
2. Run **Slide Deck Slimmer** from ScripTree.

### Run A — default slim onto a copy

| Field | Value |
|---|---|
| Remove unused custom layouts | *(checked)* |
| Also remove empty designs / slide masters | *(unchecked — opt-in)* |
| Work on a copy (leave my open deck untouched) | *(checked)* |
| Output format | `markdown` |

**Expected** — a sibling `sample_bloated_Slimmed.pptx` in which:

* the report shows `before=11`, `after=2`, **~9 unused layouts removed**
  (`removed=9`) onto the copy;
* the two used layouts (`Title Slide`, `Title and Content`) are **kept**;
* `mastersRemoved=0` (the single master still has 2 layouts, so it isn't
  empty — and the opt-in toggle was off anyway);
* the **original `sample_bloated.pptx` is untouched**;
* the saved copy is smaller than the original (unused layout parts are gone).

> Exact `removed`/`after` numbers can vary by a layout or two depending on your
> PowerPoint version's stock template and whether any layout shares a name —
> the tool is deliberately conservative and keeps anything it isn't sure is
> unused, and always keeps the last layout of a master. The headline result is
> that the handful of unused layouts are dropped and the two used ones survive.

### Run B — also drop empty designs/masters

Re-run with **Also remove empty designs / slide masters** = *checked*. On this
single-master sample it changes nothing (`mastersRemoved=0`) because the one
master still has the two used layouts. The toggle only bites on decks that have
**multiple** designs where some end up with zero layouts after the purge.

## Image / media recompression — explicitly NOT done

This tool **does not recompress images or media.** PowerPoint exposes picture
recompression only through the interactive **Compress Pictures** dialog, which
cannot be driven automatically over COM. The size reduction demonstrated here
comes **solely** from removing unused slide layouts (and, opt-in, empty
designs/masters). If you also need image recompression, do it by hand in
PowerPoint (Picture Format ▸ Compress Pictures).

## What this demonstrates

* Removing unused custom layouts to shrink a template-heavy deck.
* The conservative deletion rule: only clearly-unused layouts go; used layouts
  and the last layout of a master are always kept.
* The work-on-a-copy safety model: output lands on
  `<name>_Slimmed.pptx`, original untouched.

> Like every app in this catalog, the Slimmer is pending live verification
> against a real PowerPoint. The COM behaviours flagged for double-check are the
> layout-usage keying (master name + layout Name + Index, since RCW identity is
> unreliable) and `CustomLayout.Delete()` / `Design.Delete()` behaviour.
