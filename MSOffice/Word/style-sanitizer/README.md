# Corporate Style Sanitizer (Word)

Force the **open Word document** into a clean, consistent corporate look in one
pass: strip rogue manual (direct) formatting back to each paragraph's named
style, tidy messy whitespace, normalise body and heading fonts, and optionally
convert straight quotes to curly typographer's quotes.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

This app was co-designed with **Gemini** (feature brainstorm) and then built and
corrected by Claude — see §5 for the one COM correction made to Gemini's
proposal (the difference between `ClearFormatting()` and `Font.Reset()`).

---

## 1. What the user sees (end-user guide)

### The form fields

| Field | Meaning |
|---|---|
| **Strip manual (direct) formatting** (default ON) | Reverts every character and paragraph to its style's definition — the Ctrl+Space then Ctrl+Q equivalent. Removes rogue manual font colours, sizes, fonts, **bold/italic**, highlighting, odd indents and line-spacing, while keeping named styles, hyperlinks, table shading and list numbering. Turn OFF to keep manually-applied emphasis. |
| **Collapse repeated & trailing spaces** (default ON) | Runs of 2+ spaces → one; removes a space left before a paragraph mark. Tabs untouched. |
| **Remove blank paragraphs** (default ON) | Empty "spacer" paragraphs collapse away (a run of blank lines disappears). Table cells unaffected. |
| **Normalise fonts to the corporate pair** (default ON) | Rewrites the **Normal** style to the Body font and **Heading 1–9** to the Heading font, by editing the **style definitions** (not direct formatting). For full effect leave "Strip manual formatting" on too. |
| **Body font** (default `Calibri`) | Font for the Normal style and everything inheriting from it. |
| **Heading font** (default `Calibri Light`) | Font for the Heading 1–9 styles. Set equal to the body font for a single-font look. |
| **Style template document** (optional) | Path to a `.docx`/`.dotx` whose styles you want to adopt. When set, `CopyStylesFromTemplate` copies that file's **style definitions** into your document (headings, body, captions, …), so it matches the template. **Supersedes** the Body/Heading font fields. Pair with "Strip manual formatting" so text follows the copied styles. |
| **Convert straight quotes to curly quotes** (default OFF) | `'`→`‘’` and `"`→`“”`, opening vs closing chosen by context. Leave OFF for docs with code, paths, or inch/foot measurements. |
| **Normalise bullet glyphs** (ADVANCED, default OFF) | Standardises non-standard bullet symbols (custom Wingdings arrows, picture bullets) to the default round bullet, preserving each item's level. Numbered lists are left untouched. Opt-in because very unusual multi-level bullet styling may be simplified. |
| **Strip table cell shading** (ADVANCED, default OFF) | Removes manual cell background fills from every table (clean monochrome grids; gridlines/borders stay). Opt-in because some shading is intentional brand styling. |
| **Enforce page margins** (ADVANCED, default OFF) + **Top / Bottom / Left / Right margin (inches)** (each default 1) | Sets each side's margin of every section independently to its own value (leave all at 1 for a uniform 1″ page). Absolutely-positioned shapes keep their fixed position and may then sit outside the new margins. |
| **Work on a copy** (default ON) | SAFETY GUARD. When on, the document is first saved as `<name>_Sanitized.docx` and all changes land on that copy — the original is never modified. When off, edits are made to the open document **in memory and left unsaved** for you to review and save. |
| **Output format** | `Markdown` or `Plain text` for the result summary. |

The form is organised into four **tabs** (`sections` with `layout: "tab"`):
**Clean-up** (strip / spaces / blanks / curly quotes), **Fonts & template**
(normalise + Body/Heading fonts + style-template), **Advanced** (bullets / table
shading / margins), and **Output & safety** (work-on-copy + output format) — per
the format guide's "10+ params → prefer tab-mode sections" rule.

### What it does to your file

* **Work on a copy = ON (default):** a `<name>_Sanitized.docx` copy is created
  next to the original; all clean-up happens there; the copy is saved. Your
  original file is never touched.
* **Work on a copy = OFF:** clean-up is applied to the open document **in memory
  and left UNSAVED** — review in Word and save (or undo) yourself. The tool never
  silently overwrites the original on disk.

### Prerequisites

* The target document is **already open** in a running Word instance.
* The document has **no tracked changes** — if it does, the run is refused (see
  the safety note below). Accept/reject them first.
* For "Work on a copy", the document has been **saved at least once** (it needs a
  folder for the copy) — otherwise the run is refused.
* ScripTree's bundled `lib/combridge/combridge.exe` is present.

### Safety: tracked changes

If the document contains **any tracked changes** (`wdDoc.Revisions.Count > 0`)
the tool **refuses to run** and tells you to accept or reject them first.
Running a bulk clean-up over a document under review would otherwise turn the
whole document into a wall of revision marks. On the working copy, Track Changes
is also temporarily **disabled** so the clean-up edits do not themselves become
revisions (the user's setting is restored afterward).

---

## 2. The logic (reconstruct-the-tool spec)

A **Strategy-A shim**: a ScripTree form runs `style_sanitizer.py`; the shim bakes
the form values into a C# Roslyn script rendered from
`style_sanitizer.csx.template`, then runs it through combridge, which owns the
live COM connection to Word.

* **combridge `run-script` has no argv channel** — a `.csx` only sees the plugin
  globals (`wdApp` / `wdDoc`) and environment. Form values are baked in by
  replacing `__TOKEN__` placeholders (booleans → `true`/`false`; strings →
  C#-escaped literals).
* **combridge swallows the script's `return` value** — a clean run exits 0
  (compile=3 / throw=4 / host=5 differ). So the script prints a first-line
  **sentinel** and the shim translates it into the process exit code.

### Sentinel + status → exit-code contract

First stdout line of the `.csx`:

```
__WORDSAN__ STATUS=<code> [key=value ...]
```

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Clean-up ran; report follows (`ops=N removed=C copy=0/1`). |
| `NODOC` | 2 | No document open in Word. |
| `TRACKED` | 2 | Document has tracked changes; refused. |
| `UNSAVED` | 2 | "Work on a copy" requested but the document was never saved. |
| `BADTEMPLATE` | 2 | A style-template path was given but the file doesn't exist. |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes. |

`OK` with `ops=0` (no clean-up option ticked) is still **exit 0** — a successful
no-op, with a tip in the report.

### What the `.csx` does, in order

1. **Guards:** `wdDoc` null → `NODOC`. `wdDoc.Revisions.Count > 0` → `TRACKED`.
2. **Copy guard (if "work on a copy"):** if `wdDoc.Path == ""` (never saved) →
   `UNSAVED`. Otherwise build `copyPath = <dir>/<baseName>_Sanitized.docx` and,
   with `DisplayAlerts` off, `wdDoc.SaveAs2(copyPath, wdFormatXMLDocument)`.
   **`SaveAs2` repoints the active document at the copy**, so every edit below and
   the final `Save()` land on the copy — the original file is untouched.
3. **Disable Track Changes** on the working doc (`wdDoc.TrackRevisions = false`,
   prior value saved), and `DisplayAlerts = wdAlertsNone`, both restored in a
   `finally`. Capture `charsBefore = wdDoc.Content.Characters.Count`.
4. **Clean-up passes** (each gated by its checkbox):
   * **strip** — `Range body = wdDoc.Content; body.Font.Reset();
     body.ParagraphFormat.Reset();` Reverts ALL manual character + paragraph
     formatting to each paragraph's named **style**. `Font.Reset` = Ctrl+Space,
     `ParagraphFormat.Reset` = Ctrl+Q. Named styles, hyperlink fields, table cell
     shading and list numbering survive — only manual overrides on top go.
   * **spaces** — `ReplaceAll(" {2,}", " ", wildcards:true)` then
     `ReplaceAll(" ^p", "^p", wildcards:false)`.
   * **blanks** — `ReplaceAll("^p^p", "^p", wildcards:false)`, looped (see
     `ReplaceAll` below) so runs of 3+ blank lines fully collapse.
   * **fonts / template** — if `styleTemplate` is set,
     `wdDoc.CopyStylesFromTemplate(styleTemplate)` copies that file's style
     definitions in (it **supersedes** the font pair). Otherwise
     `wdDoc.Styles[wdStyleNormal].Font.Name = bodyFont;` and, for `i = 1..9`,
     `wdDoc.Styles[(WdBuiltinStyle)(-(i+1))].Font.Name = headingFont;` (Heading
     1–9 are the builtin styles −2…−10). Either way editing **style definitions**
     keeps the document style-driven rather than re-adding direct formatting.
     Each access is wrapped in try/catch; empty font name skips. (A missing
     template path is caught early as `BADTEMPLATE`, before the copy is made.)
   * **quotes** — save `wdApp.Options.AutoFormatAsYouTypeReplaceQuotes`, set it
     `true`, `ReplaceAll("\"","\"")` and `ReplaceAll("'","'")`, restore. With the
     option ON, replacing a straight quote *with a straight quote* makes Word
     insert the context-correct **curly** glyph; the find no longer matches the
     curly result, so the loop terminates.
   * **bullets** (ADVANCED) — `bulletTmpl =
     wdApp.ListGalleries[wdBulletGallery].ListTemplates[1]`; for each
     `Paragraph p` whose `p.Range.ListFormat.ListType` is `wdListBullet` /
     `wdListPictureBullet`, `lf.ApplyListTemplateWithLevel(bulletTmpl,
     ContinuePreviousList:true, ApplyTo:wdListApplyToSelection,
     DefaultListBehavior:wdWord10ListBehavior, ApplyLevel:lf.ListLevelNumber)`.
     Re-glyphs odd bullets to the default round bullet at the **same level**;
     numbered lists are skipped; each item is wrapped in try/catch so one
     stubborn list can't abort the pass.
   * **tables** (ADVANCED) — for each `Table t`,
     `t.Range.Shading.BackgroundPatternColor = wdColorAutomatic;
     t.Range.Shading.Texture = wdTextureNone;`. Clearing the table's **whole
     range** at once sidesteps the merged-cell exception you hit iterating
     `Rows.Cells` individually.
   * **margins** (ADVANCED) — each side converted independently with
     `wdApp.InchesToPoints(marginTop/Bottom/Left/Right)`; for each `Section s`,
     set `s.PageSetup.TopMargin/BottomMargin/LeftMargin/RightMargin` to its own
     value (all four default to 1″, i.e. a uniform page unless you change them).
5. Capture `charsAfter`; restore TrackRevisions + DisplayAlerts in `finally`.
6. **Persist only in copy-mode:** `wdDoc.Save()`. In-place mode leaves the edits
   unsaved for review.
7. **Report:** sentinel `__WORDSAN__ STATUS=OK ops={N} removed={charsBefore-After}
   copy={0/1}` then the markdown/plain-text summary (document name, the list of
   clean-up operations applied, characters removed, and the copy path / original).

### The `ReplaceAll` helper (why it loops)

```
int ReplaceAll(string findText, string replaceText, bool wildcards) {
    for (passes = 0; passes < MAX_PASSES; passes++) {
        Range rng = wdDoc.Content;            // fresh whole-doc range each pass
        Find f = rng.Find;
        f.ClearFormatting(); f.Replacement.ClearFormatting();
        f.Text = findText; f.Replacement.Text = replaceText;
        f.Forward = true; f.Wrap = wdFindStop;
        f.MatchCase = true; f.MatchWholeWord = false; f.MatchWildcards = wildcards;
        if (!f.Execute(Replace: wdReplaceAll)) break;   // no match this pass → done
    }
}
```

A single `wdReplaceAll` does **one** left-to-right scan, so newly-adjacent
matches (three blank lines collapsing into a remaining double) aren't caught.
Re-running whole-document replace-all until a pass changes nothing handles that.
`MAX_PASSES` (200) is a backstop.

### Word COM facts this relies on (see the office-com RAG)

* **`ClearFormatting()` is NOT how you clear a document's formatting.** It exists
  only on `Find` / `Replacement` and clears the **search criteria's** formatting.
  To strip a range's direct formatting use `Range.Font.Reset()` +
  `Range.ParagraphFormat.Reset()`. (This is the correction to the original
  Gemini proposal — see §5.)
* **Paragraph marks can't be inserted via `Range.Text`** — `rng.Text = "^p"`
  writes the literal characters `^` and `p`. `^p` is only interpreted inside
  Word's native Find/Replace, which is why every textual clean-up uses
  `find.Execute(Replace: wdReplaceAll)` rather than the assign-to-`Range.Text`
  pattern used by the find/replace tool.
* **`SaveAs2(path, wdFormatXMLDocument)` repoints `wdDoc`** at the new file — the
  basis of the "work on a copy" guard.
* **Heading 1–9 are builtin styles −2…−10** (`WdBuiltinStyle.wdStyleHeading1 =
  -2`); Normal is `wdStyleNormal = -1`.
* **`Document.CopyStylesFromTemplate(path)`** copies a template's/reference
  document's style definitions into the doc — the cleanest "make it match"
  mechanism. Combine with the strip pass so text follows the copied styles.
* **`AutoFormatAsYouTypeReplaceQuotes`** drives the straight→curly substitution
  during Find/Replace.
* **`Table.Range.Shading` clears a whole table at once** — clearing per-cell via
  `Rows.Cells` throws on merged cells; the table-range write avoids that trap.
* **`ApplyListTemplateWithLevel`** re-glyphs a bullet list item while preserving
  its level; pull the default bullet from
  `ListGalleries[wdBulletGallery].ListTemplates[1]` (1-based).
* **`Section.PageSetup` margins are in points** — convert with
  `wdApp.InchesToPoints(inches)`.
* **Save paths raise modal dialogs that hang a hidden COM-launched Word** — set
  `wdApp.DisplayAlerts = wdAlertsNone` around them.

---

## 3. Files in this app

| File | Role |
|---|---|
| `style-sanitizer.scriptree` | The form (17 params across 4 tabs — 5 core + style-template + 3 advanced opt-in passes + 4 per-side margins + work-on-copy + output; a broom-sweeping-sparkles PNG icon embedded — the clean/sweep convention; this tool MUTATES the doc). |
| `style-sanitizer.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config. |
| `style_sanitizer.py` | The Strategy-A shim. |
| `style_sanitizer.csx.template` | The Roslyn template with `__TOKEN__` placeholders. |
| `examples/` | A generator + a deliberately-messy `sample_messy.docx` + a README. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
style_sanitizer.py
  [--strip-formatting]            (flag; default ON in the form)
  [--collapse-spaces]             (flag; default ON)
  [--remove-blank-paragraphs]     (flag; default ON)
  [--normalize-fonts]             (flag; default ON)
  --body-font <str>               (default "Calibri")
  --heading-font <str>            (default "Calibri Light")
  --style-template <path>         (optional; supersedes the font fields)
  [--smart-quotes]                (flag; default OFF)
  [--list-normalize-bullets]      (flag; ADVANCED, default OFF)
  [--strip-table-shading]         (flag; ADVANCED, default OFF)
  [--enforce-margins]             (flag; ADVANCED, default OFF)
  --margin-top <float>            (default 1.0)   } per-side margins, inches;
  --margin-bottom <float>         (default 1.0)   } each baked as a C# double
  --margin-left <float>           (default 1.0)   } literal
  --margin-right <float>          (default 1.0)   }
  [--work-on-copy]                (flag; default ON)
  --output-format markdown|text
```

The booleans use the `{id?--flag}` conditional form; `--body-font` /
`--heading-font` use `["--flag","{id}"]` token groups with `no_split: true` so a
font name with spaces ("Calibri Light") stays a single token. The four
`--margin-*` values are `number` params; the shim bakes each via `repr(float(x))`
so the rendered C# `double` always carries a decimal point.

---

## 4. Editing / maintenance notes

* **The "work on a copy" guard and the tracked-changes refusal are the headline
  safety properties** — don't let an edit make the in-place path overwrite the
  original on disk, and don't let the tool run over a revision-bearing document.
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree validate <path>`.
* **Offline render-check:** render the template with sample values (a font name
  containing a quote and a backslash, smart-quotes on) and grep for `__[A-Z_]+__`
  — only the `__WORDSAN__` sentinel should match; anything else is an unfilled
  placeholder.
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path.

---

## 5. Design provenance & the Gemini correction

Gemini proposed the must-have feature set (double-space collapse, empty-paragraph
purge, direct-formatting strip, two-tier font normaliser, smart quotes) and the
tracked-changes safety abort, and correctly scoped out the risky items
(list/bullet normalisation, table-shading stripping, page-margin enforcing) as
COM-fragile. Two changes were made when building:

1. **COM correction.** Gemini's implementation column listed
   `Paragraph.Range.ClearFormatting()` for the direct-formatting strip. That is
   not a real document operation — `ClearFormatting()` only clears **search
   criteria** on `Find`/`Replacement`. The strip uses `Range.Font.Reset()` +
   `Range.ParagraphFormat.Reset()` instead.
2. **Font-normaliser refinement.** Rather than looping paragraphs and applying a
   direct font (which would re-introduce the very direct formatting the strip
   removes), the normaliser edits the **style definitions** (Normal + Heading
   1–9). The document stays clean and style-driven.

Treat any tool/LLM-sourced implementation detail as untrusted until verified —
that is exactly how the `ClearFormatting` error was caught.
