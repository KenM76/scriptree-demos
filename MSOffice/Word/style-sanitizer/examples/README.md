# Example — Corporate Style Sanitizer

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates both samples. |
| `sample_messy.docx` | A deliberately messy document: rogue manual fonts/colours/sizes, double & triple spaces, trailing spaces, blank spacer paragraphs, straight quotes, three real `Heading 1/2` headings, a **real hyperlink**, and a **shaded table**. |
| `sample_template.docx` | A clean reference document with distinct **style definitions** (Georgia 11pt body; navy Cambria headings). Point the **Style template document** field at this file to make the messy doc adopt these styles. |

### What's seeded, and why

| Mess | Where | Which option fixes it |
|---|---|---|
| Manual font/size/colour/bold overrides | Comic Sans/Arial/Georgia/Times/Courier runs on headings & body | **Strip manual formatting** → reverts to the style |
| `double  spaces`, `triple   spaces`, trailing spaces | every body paragraph (20 double-space runs) | **Collapse repeated & trailing spaces** |
| 6 empty "spacer" paragraphs | between sections | **Remove blank paragraphs** |
| Straight `'` and `"` quotes | 6 of them in body text | **Convert straight quotes to curly** (off by default) |
| Headings on `Heading 1`/`Heading 2` styles | three of them | **Normalise fonts** → rewrites the heading styles |

Two things are there to prove they **survive** (must be preserved):

* A **real hyperlink** (`Corporate Standards` → example.com) — the link must
  still work after sanitising; only its manual blue/underline reverts when
  formatting is stripped (the URL field is untouched).
* A **shaded table** (blue header row, grey body) — the cell background fill must
  remain. Paragraph-format reset does not touch cell shading.

## How to try it

1. Open `sample_messy.docx` in Word **and save it once** (it's freshly
   generated; the copy-mode default needs a saved file to put the copy beside).
2. Run **Corporate Style Sanitizer** from ScripTree.

Form values (the defaults, plus turning on curly quotes to see that pass):

| Field | Value |
|---|---|
| Strip manual (direct) formatting | *(checked)* |
| Collapse repeated & trailing spaces | *(checked)* |
| Remove blank paragraphs | *(checked)* |
| Normalise fonts to the corporate pair | *(checked)* |
| Body font | `Calibri` |
| Heading font | `Calibri Light` |
| Convert straight quotes to curly quotes | *(checked)* |
| Work on a copy | *(checked)* |
| Output format | `markdown` |

**Expected result** — a sibling `sample_messy_Sanitized.docx` in which:

* every Comic Sans / Georgia / Courier override is gone — body text is **Calibri**,
  headings are **Calibri Light**, all via the styles;
* the 20 double/triple-space runs are single spaces; trailing spaces are gone;
* the 6 blank paragraphs are collapsed away;
* `'single'` / `"double"` straight quotes are now curly `‘single’` / `“double”`;
* the **hyperlink still works** (now styled like body text, but the URL is
  intact and clickable);
* the **table shading is unchanged** (blue header, grey body row);
* your original `sample_messy.docx` is **untouched**.

The report's **"Characters removed"** count reflects the whitespace and blank
lines deleted (font/quote changes don't change the count).

### Variations to try

* **Keep emphasis:** untick **Strip manual formatting**. Now bold survives, but
  so do the rogue fonts/colours — and because the fonts weren't stripped,
  **Normalise fonts** can't fully take on text that has a hard-coded font (the
  style change only reaches text that follows its style). This shows why strip +
  normalise are designed to run together.
* **Single-font look:** set **Heading font** = `Calibri` too — headings and body
  share one typeface.
* **In place:** untick **Work on a copy** — the open document is cleaned in
  memory and left **unsaved** for you to review (Ctrl+Z still undoes it).

### The ADVANCED opt-in passes (all OFF by default)

The sample is also seeded for the three advanced options. Tick them (keep
**Work on a copy** on) to see:

* **Normalise bullet glyphs** → the three **Action Items** bullets are
  re-glyphed to the standard round bullet at their level. (They start as
  standard bullets, so the visible change is nil but the report counts 3 items;
  to see a *non-standard* bullet normalised, change one to a Wingdings arrow in
  Word first — see the note in `make_example.py`.)
* **Strip table cell shading** → the Summary Table's blue header and grey body
  fills are cleared to a clean monochrome grid (the borders stay). Note this is
  the *opposite* of the default run above, where shading is preserved — shading
  is only removed when you tick this advanced option.
* **Enforce page margins** → set the four per-side fields (e.g. Top `1`, Bottom
  `1`, Left `1.25`, Right `0.75`) and the document's `0.5"` margins become those
  values on every section. Leave all four at `1` for a uniform 1″ page.

### Match a template (the `Style template document` option)

Instead of typing fonts, point the sanitizer at `sample_template.docx`:

| Field | Value |
|---|---|
| Strip manual (direct) formatting | *(checked)* |
| Style template document | *(browse to `sample_template.docx`)* |
| Work on a copy | *(checked)* |

**Expected** — in `sample_messy_Sanitized.docx`, the document's style definitions
are replaced by the template's (`CopyStylesFromTemplate`): body text becomes
**Georgia 11pt**, headings become **navy Cambria**. Because the strip pass ran
first, the text actually follows the copied styles. The **Style template
document** field **supersedes** the Body/Heading font fields — those are ignored
when a template is supplied.

### The tracked-changes refusal

Turn on **Review → Track Changes** in Word, type a word, then run the tool. It
**refuses** with a `TRACKED` message asking you to accept/reject first — a bulk
clean-up would otherwise drown the document in revision marks.

## What this demonstrates

* One-pass "force into a clean corporate standard": strip direct formatting,
  fix whitespace, normalise fonts, curl quotes.
* Editing **style definitions** (not direct formatting) for the font pass, so the
  result stays clean and style-driven.
* The preserved-asset guarantees (named styles, hyperlinks, table shading, list
  numbering).
* The work-on-a-copy default and the tracked-changes safety refusal.

> Like every app in this catalog, the sanitizer is pending live verification
> against a real Word. The scenario above describes the exact behaviour to
> expect from a messy document with these features.
