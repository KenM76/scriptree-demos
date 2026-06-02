# Global Typography Enforcer (PowerPoint)

Standardise the fonts in a PowerPoint deck — either swap **one** named
font for another (brand compliance) or collapse **every** text font to a
single typeface (taming a "frankendeck") — and write the result to a
**sibling copy**. The deck you have open is never modified.

This document is written to the project's reconstruct-from-docs bar: a
competent engineer or LLM should be able to rebuild the tool from this
README alone. It explains the *logic*, the *contracts*, the *safety
model*, and the *edge cases* — not just the surface behaviour.

---

## 1. What it does, and when to use it

You have a `.pptx` open in PowerPoint. You run this tool from ScripTree
and pick a mode:

* **Replace a single specific font** *(default)* — every text run set in
  the font you name (e.g. `Calibri`) is re-set to your target font (e.g.
  `Aptos`). Every other font is left exactly as it was. This is the
  precision scalpel for "our brand standard is Aptos, but Calibri keeps
  sneaking in."
* **Replace ALL text fonts** — every text run in the deck is set to the
  target font. This is the blanket hammer for a deck assembled from many
  sources, where you just want one consistent typeface.

In **both** modes, a hard-coded blocklist of **symbol/icon fonts** is
always bypassed: `Wingdings`, `Wingdings 2`, `Wingdings 3`, `Webdings`,
`Symbol`, `Marlett`. Those fonts encode pictographs by codepoint —
reflowing a Wingdings run to a text font would turn ✆ ➜ ✓ into the
letters that happen to share those codepoints. Skipping them keeps icons
intact.

By default the change is **also applied to Slide Masters & Custom
Layouts** (`Enforce on Slide Masters & Layouts`, on by default). This
matters: if you only restyle the visible slides, any *new* slide you add
afterwards inherits the *old* font from the master, and the deck appears
to "revert." Turn this off only if you deliberately want the masters
untouched.

**Out of scope in this version** (deferred, not silently mishandled):

* **SmartArt** graphics — PowerPoint's SmartArt text API
  (`Shape.SmartArt.AllNodes` via `TextFrame2`) is erratic over COM across
  Microsoft 365 patch levels, so SmartArt shapes are **skipped and
  counted**, never touched. The report tells you how many were skipped so
  you can fix them by hand.
* **Theme font schemes (writing them on *this* deck)** — in the manual and
  template-`fonts` paths the tool sets `Font.Name` on each **run** rather
  than rewriting the deck's own theme major/minor scheme; runs that inherit
  from the theme are still changed directly. (To adopt a *template's* whole
  theme — including its scheme — use the template-`theme` mode, which calls
  `ApplyTemplate`; see §3.1. And note the template-`fonts` mode *reads* a
  template deck's theme fonts.)

### Requirements

* The target deck is **already open** in a running PowerPoint instance.
* The deck has been **saved at least once** (it needs a folder on disk to
  put the copy in).
* `combridge` is available — bundled with ScripTree at
  `lib/combridge/combridge.exe`. The shim finds it by walking up the
  directory tree (see §6); no absolute path is baked in.

---

## 2. The output: a sibling copy, original untouched

The tool writes `<DeckName>_Restyled.pptx` next to the original (same
folder). **Your open deck is provably untouched** — see the safety model
in §5. Open the `_Restyled` copy to review the result; if you're happy,
keep it; if not, delete it and your original is exactly as it was.

The output pane shows a summary (Markdown or plain text, your choice):
how many text runs were changed, how many slides / masters / layouts
were swept, and how many SmartArt shapes were skipped.

---

## 3. The form (7 parameters, grouped into 3 sections)

The form is grouped into three collapsible sections:

* **What to change** — `replacement_mode`, `source_font`, `target_font`
  (the manual font-swap fields).
* **Template** — `font_template`, `template_mode` (the "read styles from a
  template deck" fields; see §3.1).
* **Scope & output** — `include_masters`, `output_format`.

(The sections use `layout: "collapse"`, so each renders as a collapsible
group box.)

| id | label | type / widget | default | section | meaning |
|---|---|---|---|---|---|
| `replacement_mode` | Replacement mode | `enum` / `radio` | `specific` | What to change | `specific` = swap one named font; `all` = set every run to the target. A radio (not a checkbox) because it gates whether `source_font` is read. **Ignored when a template deck is set.** |
| `source_font` | Target font to remove | `string` / `text` | `""` | What to change | **Only used in `specific` mode.** Exact font name to replace (case-insensitive match). Leave blank for `all` mode. **Ignored when a template deck is set.** |
| `target_font` | New font to apply | `string` / `text` | `""` | What to change | **Required in manual mode.** The font every affected run is set to. **Ignored (not required) when a template deck is set.** |
| `font_template` | Reference template deck (optional) | `path` / `file` | `""` | Template | OPTIONAL reference `.pptx`/`.potx` to pull styling from. When set, the manual fields above are ignored. The template deck is opened **read-only** and never modified. See §3.1. |
| `template_mode` | How to use the template | `enum` / `radio` | `fonts` | Template | Only relevant when `font_template` is set. `fonts` = read the template's theme heading/body fonts and sweep; `theme` = apply the template's whole theme via `ApplyTemplate`. See §3.1. |
| `include_masters` | Enforce on Slide Masters & Layouts | `boolean` / `checkbox` | `true` | Scope & output | Also rewrite fonts on Slide Masters + Custom Layouts so new slides don't revert. (Applies to the manual sweep and the template-`fonts` sweep; the template-`theme` path applies the whole theme regardless.) |
| `output_format` | Output format | `enum` / `dropdown` | `markdown` | Scope & output | `markdown` or `text` rendering of the summary. |

Type the font names **exactly** as they appear in the Windows / PowerPoint
font list (e.g. `Segoe UI`, `Times New Roman`). Matching on `source_font`
is case-insensitive but otherwise exact — it is not a fuzzy or
family-prefix match.

### 3.1 Reading styles from a template deck

If you fill in **`font_template`** (a reference `.pptx`/`.potx` — a corporate
template or any well-styled deck), the tool **ignores** the manual
`replacement_mode` / `source_font` / `target_font` fields and instead pulls
styling from that deck. The `template_mode` radio chooses how:

* **`fonts` — "Use the template's theme fonts (headings + body)"** *(default)*.
  The tool opens the template deck **read-only and headless**, reads its theme
  **MAJOR** (heading) and **MINOR** (body) Latin fonts
  (`SlideMaster.Theme.ThemeFontScheme.MajorFont/MinorFont[msoThemeLatin]`),
  closes it, then sweeps the working copy: every **title / subtitle**
  placeholder's text is set to the heading font, and **all other text** is set
  to the body font. Symbol/icon fonts are still preserved. This is a pure
  **typography** change.

  If the template's theme fonts can't be read (or come back blank), the tool
  refuses with `STATUS=NOTHEME` (exit 2) and tells you to try the full-theme
  mode or enter fonts manually. **Caveat:** this mode reads the template's
  *theme font scheme*. A deck that has no proper theme scheme (e.g. one built
  by python-pptx for the bundled example) may report `NOTHEME`; a real
  corporate `.potx` carries the scheme and works as intended. See the example
  README for the full honesty note.

* **`theme` — "Apply the full template theme (fonts + colours + masters)"**.
  After `SaveCopyAs` + headless-open of the working copy, the tool calls
  `work.ApplyTemplate(font_template)`. This applies the template deck's
  **entire theme** — fonts AND colours AND master/layout design — to the copy.
  This is **broader than typography** (it is not a pure font sweep); it is the
  most reliable way to make a deck adopt a corporate template's look, and it
  works regardless of how the template was authored. The per-run sweep is
  skipped in this mode.

In **both** template modes the `NOTARGET` / `NOSOURCE` guards do **not** apply
— no manual fonts are required. The only template-specific refusals are
`BADTEMPLATE` (the named file doesn't exist) and, for `fonts` mode,
`NOTHEME`.

### argv contract (argument_template → shim)

```
font_enforcer.py
["--replacement-mode", "{replacement_mode}"]
["--source-font",      "{source_font}"]      # token group DROPS when blank
["--target-font",      "{target_font}"]
"{include_masters?--include-masters}"        # emits the flag only when checked
["--font-template",    "{font_template}"]    # token group DROPS when blank
["--template-mode",    "{template_mode}"]
["--output-format",    "{output_format}"]
```

The `["--source-font", "{source_font}"]` token **group** disappears
entirely when `source_font` is empty (ScripTree's drop-on-empty rule), so
in `all` mode no `--source-font` reaches the shim — `argparse` then uses
its default `""`. Likewise `["--font-template", "{font_template}"]` drops
when the template field is blank, so the shim's `--font-template` defaults
to `""` (the manual path). The `{include_masters?--include-masters}`
conditional emits the flag string only when the box is checked.

---

## 4. Exit-code contract (sentinel → exit code)

combridge's `run-script` **ignores the C# `return` value** (it exits 0 on
any clean run) and offers **no argv channel**. So the generated `.csx`
writes a **first-line sentinel** that the Python shim parses, and the
**shim** owns the exit code ScripTree finally sees:

| Sentinel first line | Exit | Meaning |
|---|---|---|
| `__PPTFONT__ STATUS=NODECK` | **2** | No presentation is open. |
| `__PPTFONT__ STATUS=UNSAVED` | **2** | Deck never saved → no folder for the copy. |
| `__PPTFONT__ STATUS=NOTARGET` | **2** | *(manual path only)* No target font was given. |
| `__PPTFONT__ STATUS=NOSOURCE` | **2** | *(manual path only)* `specific` mode but no source font given. |
| `__PPTFONT__ STATUS=BADTEMPLATE` | **2** | *(template path)* The named template deck doesn't exist (checked early, before any copy is written). |
| `__PPTFONT__ STATUS=NOTHEME` | **2** | *(template `fonts` path)* The template's theme major/minor fonts couldn't be read or were blank. |
| `__PPTFONT__ STATUS=OK path=… changed=… slides=… masters=… layouts=… smartart_skipped=…` | **0** | Success. `path` is `manual`, `template-theme`, or `template-fonts`. |
| *(combridge's own codes)* | **3 / 4 / 5** | Compile error / script threw / host (no PowerPoint) — passed through verbatim. |

`changed=0` is **still exit 0** — "no run used the source font" (or "every
run was already the target font") is a legitimate no-op success, not an
error. The guarded refusal statuses are the only `STATUS=` values
that map to exit 2. In **template** mode `NOTARGET`/`NOSOURCE` are not raised
(no manual fonts required); the template-specific refusals are `BADTEMPLATE`
and `NOTHEME`.

---

## 5. Safety model — why the open deck is never touched

PowerPoint's `Presentation.SaveCopyAs` is the linchpin. Unlike Word's
`SaveAs2` (which *repoints* the document object at the new file),
`SaveCopyAs` writes a copy to disk and **leaves the active presentation
pointed at the original**. So the sequence is:

1. `pptPres.SaveCopyAs(copyPath, ppSaveAsOpenXMLPresentation)` — write
   `<name>_Restyled.pptx`. `pptPres` (the user's deck) is unchanged and
   still open.
2. `pptApp.Presentations.Open(copyPath, ReadOnly:msoFalse,
   Untitled:msoFalse, WithWindow:msoFalse)` — open the **copy headless**
   (no window). All font rewriting happens on this `work` presentation.
3. `work.Save()` then `work.Close()` (in a `finally`).

Because every `Font.Name` write targets `work` (the copy), the deck the
user is looking at cannot be modified, even if the script throws midway.

`pptApp.DisplayAlerts = PpAlertLevel.ppAlertsNone` is set around the
whole block and restored in the `finally`. A hidden/headless COM-attached
PowerPoint **hangs invisibly** on any modal prompt (overwrite, repair,
font-substitution), so alerts must be suppressed — same hazard documented
for COM-launched Excel.

Guards before any mutation:

* `pptPres is null` → `NODECK` (nothing open).
* `pptPres.Path == ""` → `UNSAVED` (never-saved deck has no folder).
* *(manual path only)* `target_font` blank → `NOTARGET`.
* *(manual path only)* `specific` mode + blank `source_font` → `NOSOURCE`.
* *(template path)* `font_template` doesn't exist → `BADTEMPLATE`, checked
  **early** (before any copy is written, so a bad path leaves no stray
  `_Restyled` file).
* *(template `fonts` path)* theme major/minor fonts unreadable/blank →
  `NOTHEME`, also checked **before** the copy is written.

The template deck (when supplied) is opened **read-only** and headless
purely to read its theme font names, then closed immediately — it is never
modified.

### 5.1 COM facts the template paths rely on

These are the PowerPoint COM calls the template feature uses. They are
flagged for live verification (the interop signatures/enum values are the
parts most worth confirming against a real PowerPoint):

* **Open a deck read-only headless:**
  `pptApp.Presentations.Open(path, ReadOnly:msoTrue, Untitled:msoFalse,
  WithWindow:msoFalse)` — used for the template deck in `fonts` mode (we only
  read its theme, never edit it).
* **Read theme heading/body fonts:**
  `tpl.SlideMaster.Theme.ThemeFontScheme.MajorFont[(MsoFontLanguageIndex)1].Name`
  for the heading font and `.MinorFont[(MsoFontLanguageIndex)1].Name` for the
  body font. `msoThemeLatin == 1` (the Latin script slot of
  `MsoFontLanguageIndex`). Wrapped in try/catch → `NOTHEME` on any failure.
* **Apply a whole theme:** `work.ApplyTemplate(font_template)` — copies the
  template deck's entire theme (fonts + colours + masters) onto the working
  copy (the `theme` mode).
* **Placeholder title detection:** a shape is a title when
  `sh.Type == MsoShapeType.msoPlaceholder` **and** `sh.PlaceholderFormat.Type`
  is one of `ppPlaceholderTitle` (**13**), `ppPlaceholderCenterTitle` (**11**),
  or `ppPlaceholderSubtitle` (**4**). Accessing `.PlaceholderFormat` on a
  non-placeholder shape **throws**, so the access is wrapped in try/catch and
  any throw means "not a title" → route to the body font.

---

## 6. How the `.csx` rewrites fonts (the core logic)

### 6.1 Per-run, not per-TextRange

Setting `Font.Name` on a whole `TextRange` would clobber symbol runs and
(in `specific` mode) non-matching runs. A **run** is a maximal span of
uniform character formatting, so each run has exactly one `Font.Name` we
can test and conditionally rewrite. Per-run is therefore the correct
granularity.

### 6.2 Indexing-robust run walk

We avoid any assumption about a run-collection index base. PowerPoint's
`TextRange.Runs(Start, Length)` returns the **whole** run(s) overlapping
the character span `[Start, Start+Length)`, with `Start` **1-based and
relative to the TextRange**. So we walk:

```csharp
int pos = 1;                       // 1-based, relative to this TextRange
while (pos <= tr.Length) {
    TextRange run = tr.Runs(pos, 1);   // the whole run overlapping char pos
    int rlen = run.Length;
    if (rlen <= 0) break;              // guard against a 0-length run
    string fn = run.Font.Name;
    // ... apply the rule (below) ...
    pos += rlen;                       // advance to the next run's first char
}
```

Because `Runs(pos, 1)` returns the *entire* run (not clipped to one
char), `pos += run.Length` lands exactly on the next run's first
character. A `guard` counter caps the loop at `tr.Length` iterations as a
belt-and-suspenders against a pathological zero-length run.

> **Live-run verification note.** The 1-based, range-relative semantics of
> `Runs(Start, Length)` is the one behaviour to confirm against a real
> PowerPoint session. The walk is written to be robust regardless (it
> derives each step from the run's own `Length`), and it only ever edits
> the **copy**, so a mis-step damages a throwaway file, never the
> original. See the office-com RAG lesson `ppt_font_sweep_per_run.md`.

### 6.3 The rule applied to each run

```
fn = run.Font.Name
if fn is a symbol font (Wingdings / Wingdings 2 / Wingdings 3 /
                        Webdings / Symbol / Marlett, case-insensitive):
    skip            # always preserve icon fonts
elif mode == "specific":
    if fn == sourceFont (case-insensitive):  run.Font.Name = targetFont; changed++
else:  # mode == "all"
    if fn != targetFont (case-insensitive):  run.Font.Name = targetFont; changed++
```

Every run access is wrapped in try/catch — a single run we can't read or
set is skipped, never fatal.

### 6.4 Shape traversal (recursive)

`SweepShape(shape)` dispatches by shape kind, each branch in its own
try/catch so one bad shape never aborts the sweep:

* **Group** (`Type == msoGroup`) → recurse into `shape.GroupItems`
  (a group has no `TextFrame` of its own).
* **Table** (`HasTable == msoTrue`) → for each `Table.Cell(r, c)`, sweep
  `cell.Shape.TextFrame`. (`Rows.Count` × `Columns.Count`, both 1-based.)
* **SmartArt** (`Type == msoSmartArt`) → `smartArtSkipped++`, return
  (deferred — see §1).
* **Plain text shape** (`HasTextFrame == msoTrue`) → sweep
  `shape.TextFrame` directly. Covers titles, bodies, text boxes,
  autoshapes, placeholders.

`SweepShapes(shapes)` just iterates and calls `SweepShape` on each.

### 6.5 What gets swept

* **Every normal slide:** `SweepShapes(slide.Shapes)` for each slide in
  `work.Slides` (`slidesVisited++`).
* **Masters + layouts (when `include_masters`):** for each `Design d` in
  `work.Designs`, sweep `d.SlideMaster.Shapes` (`mastersVisited++`) and
  every `d.SlideMaster.CustomLayouts` (`layoutsVisited++`). Iterating
  `Designs` covers **multi-master** decks. Notes pages are intentionally
  **not** swept (out of scope — they aren't presented).

§6.1–6.5 describe the **manual** path. The template paths use **separate
helpers** so the manual sweep is never altered.

### 6.6 The template paths (kept separate from the manual sweep)

The template feature is implemented with its own helpers, deliberately
parallel to `SweepShape`/`SweepTextFrame` rather than entangled with them,
so the original manual behaviour is untouched:

* **`SweepTextFrameTo(TextFrame tf, string font)`** — the same indexing-robust
  run-walk as §6.2, but it sets *every* non-symbol run to `font` (skipping a
  run already in `font`). Reuses the `IsSymbolFont` blocklist.
* **`IsTitlePlaceholder(Shape sh)`** — true when the shape is a
  `msoPlaceholder` whose `PlaceholderFormat.Type` is `ppPlaceholderTitle` (13),
  `ppPlaceholderCenterTitle` (11), or `ppPlaceholderSubtitle` (4). Wrapped in
  try/catch (a non-placeholder shape throws on `.PlaceholderFormat`).
* **`SweepShapeTemplate(Shape sh, string bodyFont, string headingFont)`** —
  parallel to `SweepShape`: a **title/subtitle** placeholder → `headingFont`,
  everything else → `bodyFont`; recurses **groups** (members → `bodyFont`) and
  **tables** (each cell → `bodyFont`); SmartArt is counted+skipped as before.

**`fonts` mode** (`template-fonts`): read the template's theme major/minor
fonts (§5.1) **before** writing any copy (so an unreadable theme fails as
`NOTHEME` with no stray file), then run `SweepShapesTemplate(slide.Shapes,
minorFont, majorFont)` over every slide, plus masters/layouts when
`include_masters` (placeholder title-style → heading, else body). Counters and
the report's `path=template-fonts` reflect this.

**`theme` mode** (`template-theme`): after `SaveCopyAs` + headless-open, call
`work.ApplyTemplate(font_template)` and skip the per-run sweep entirely. The
report's `path=template-theme` and a note explain that the whole theme
(fonts + colours + masters) was applied — broader than typography.

---

## 7. Files in this app

| File | Role |
|---|---|
| `font_enforcer.py` | Strategy-A shim. Parses argv, escapes values, renders the `.csx` from the template, runs combridge, parses the sentinel, owns the exit code. |
| `font_enforcer.csx.template` | Roslyn script template. Token placeholders (`__REPLACEMENT_MODE__`, `__SOURCE_FONT__`, `__TARGET_FONT__`, `__INCLUDE_MASTERS__`, `__OUTPUT_FORMAT__`, `__FONT_TEMPLATE__`, `__TEMPLATE_MODE__`) are filled by the shim. Contains all the COM logic above (manual sweep + the two template paths). |
| `typography-enforcer.scriptree` | The ScripTree form (7 params in 3 sections). Carries the embedded `edit` pencil icon (mutator convention). |
| `typography-enforcer.scriptree.configs.json` | Sidecar with the end-user `standalone` config (IDE chrome hidden, error/success popups on). |
| `README.md` | This document. |

### Token-substitution + escaping contract

The shim's `render_csx` replaces the seven `__TOKEN__`s. String values
(`replacement_mode`, `source_font`, `target_font`, `output_format`,
`font_template`, `template_mode`) go through `csharp_literal` — backslash
escaped **first**, then `"`, `\r`, `\n`, `\t` — so a font name or a Windows
path containing a quote or backslash can't break out of the C# string
literal (e.g. `C:\Temp\my "brand".potx` renders safely). The boolean
(`include_masters`) is rendered as the bare C# literal `true`/`false` via
`cs_bool`. After substitution the only `__…__` token remaining in the
generated `.csx` is the `__PPTFONT__` sentinel (verified by an offline
render-check in both template modes with an awkward quote+backslash path).

---

## 8. Maintenance notes

* **Adding a font to the symbol blocklist:** edit the `symbolFonts`
  `HashSet` in `font_enforcer.csx.template`. It's case-insensitive
  (`StringComparer.OrdinalIgnoreCase`).
* **Enabling SmartArt later:** replace the `smartArtSkipped++` branch in
  `SweepShape` with a traversal of `shape.SmartArt.AllNodes`, reading each
  node's `TextFrame2.TextRange` runs. Wrap heavily in try/catch and
  test against multiple M365 patch levels — this is why it's deferred.
* **Changing the copy suffix:** the `_Restyled` suffix is built in the
  `.csx` (`baseName + "_Restyled.pptx"`). Keep the sibling-copy +
  headless-reopen safety model intact if you touch it.
* **Template paths:** the template feature lives in *separate* helpers
  (`SweepTextFrameTo`, `IsTitlePlaceholder`, `SweepShapeTemplate`) and the
  `ReadTemplateThemeFonts` reader (§5.1, §6.6). Do not fold them back into the
  manual `SweepShape`/`SweepTextFrame` — keeping them apart is what guarantees
  the original manual behaviour stays unchanged. The placeholder-type enum
  values (13/11/4) and the `ThemeFontScheme` access path are the parts to
  re-verify if PowerPoint interop changes.
* **Validate after any form edit:** from `D:\Dev\ScripTree`, run
  `python -m scriptree validate <path>` (PowerShell:
  `Set-Location D:\Dev\ScripTree; python -m scriptree validate "<path>"`).
