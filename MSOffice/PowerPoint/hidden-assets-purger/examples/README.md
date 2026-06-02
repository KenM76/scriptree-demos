# Example — Hidden Assets & Notes Purger

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates the deck. |
| `sample_deck.pptx` | A three-slide deck seeded with notes, metadata, and one hidden slide. |

### What's seeded, and why

The purger strips speaker notes, comments, author/document metadata, and
(opt-in) hidden slides, writing the result to a sibling `_Sanitized` copy.
The sample carries one of each so you can confirm it's gone in the copy:

| Asset | Where | How to confirm removal |
|---|---|---|
| **Speaker notes** | on all 3 slides | open the copy's notes pane — the `NOTE:` text is gone |
| **Author / metadata** | core properties: author `Jane Author`, last-modified-by `Reviewer Bob`, keywords `internal, draft, do-not-share`, category `Confidential`, plus a core comment | File ▸ Info ▸ Properties on the copy shows them cleared |
| **Hidden slide** | slide 3, "Appendix — internal numbers", marked hidden (`show="0"`) | with the opt-in ON it's gone from the copy; with it OFF it stays (but hidden) |

### Comments — a deliberate gap

PowerPoint review **comments** can't be authored cleanly through the library
that generates this sample (they live in separate comment/author parts), so
the deck ships **without** comments. To exercise **Remove comments**: open the
deck in PowerPoint, add one by hand (**Review ▸ New Comment**), save, then run
the tool and check the copy. The notes / metadata / hidden-slide removals are
fully demonstrated by the generated file as-is.

## How to try it

1. Open `sample_deck.pptx` in PowerPoint.
2. Run **Hidden Assets & Notes Purger** from ScripTree.

### Run A — default sweep (keep hidden slides)

| Field | Value |
|---|---|
| Remove speaker notes | *(checked)* |
| Remove comments | *(checked)* |
| Strip author & document metadata | *(checked)* |
| Permanently delete hidden slides | *(unchecked — opt-in, off by default)* |
| Output format | `markdown` |

**Expected** — a sibling `sample_deck_Sanitized.pptx` in which:

* all speaker notes are **removed** (all 3 slides);
* author / last-modified-by / keywords / category / core comment are
  **cleared**;
* the hidden slide 3 is **still present** (still hidden) — because the delete
  toggle was off;
* the original deck is untouched.

### Run B — also drop hidden slides

Re-run with **Permanently delete hidden slides** = *checked*.

**Expected** — same as Run A, but the copy now has **2 slides**: the hidden
"Appendix — internal numbers" slide is gone entirely. (This is the only
destructive option, which is why it's off by default and opt-in.)

## What this demonstrates

* Stripping speaker notes, document metadata, and comments in one pass.
* The opt-in, off-by-default hidden-slide deletion (the one irreversible
  action), kept separate from the always-safe metadata/notes sweep.
* The work-on-a-copy safety model: output lands on `<name>_Sanitized.pptx`,
  original untouched.

> Like every app in this catalog, the purger is pending live verification
> against a real PowerPoint. Note that **modern threaded comments** may use
> parts the classic comment-removal path doesn't reach; the scenario above
> targets classic comments and the always-present notes/metadata/hidden-slide
> assets.
