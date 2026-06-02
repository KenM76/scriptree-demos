# Slide Deck Slimmer (PowerPoint)

Reduce the **saved file size** of the open PowerPoint deck by removing **unused
custom slide layouts** — and, optionally, any design/slide-master left with no
layouts in use. Template-heavy decks carry many layouts no slide uses, and each
unused layout can drag along its own background images and logos that bloat the
file. By default the tool works on a sibling `<name>_Slimmed.pptx` **copy**,
leaving your open deck untouched.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

> **CLAIM HONESTY — read this.** This tool does **NOT** recompress images or
> media. PowerPoint exposes picture recompression only through an interactive
> **Compress Pictures** dialog that cannot be driven fire-and-forget over COM.
> The **only** size reduction this tool produces comes from removing unused
> slide layouts (and, opt-in, empty designs/masters). Do not describe it as an
> image/media compressor anywhere.

---

## 1. What the user sees (end-user guide)

### The form fields

| Field | Default | Meaning |
|---|---|---|
| **Remove unused custom layouts** | ON | Delete every custom slide layout that no slide uses. This is the main slimming action. Conservative — see §2. |
| **Also remove empty designs / slide masters** | OFF (opt-in) | After the layout purge, delete any design (slide master) that has zero layouts left **and** that no slide uses. More aggressive; the last design is always kept. |
| **Work on a copy (leave my open deck untouched)** | ON | ON: save a sibling `<name>_Slimmed.pptx` and edit only that copy (requires the deck to have been saved once). OFF: slim the open deck in place, left open and UNSAVED for you to review. |
| **Output format** | `Markdown` | `Markdown` (paste into a ticket) or `Plain text` (aligned list). |

### What it does to your deck

* **Copy mode (default):** nothing to the original. The tool saves a copy named
  `<name>_Slimmed.pptx` in the **same folder**, performs every removal on that
  copy, then leaves the copy saved and closed. You open the copy yourself.
* **In-place mode (`Work on a copy` off):** the open deck is edited in memory
  and **left open and UNSAVED**. Review the result in PowerPoint and press
  Ctrl+S to keep it, or close without saving to discard. The tool never calls
  `Save()` on your original file.

### What it explicitly does NOT do

* **No image or media recompression** (see the claim-honesty box above).
* **No content changes** — text, pictures, embedded objects, and the layouts
  your slides actually use are all left exactly as they were.

### Prerequisites

* PowerPoint is **running with the target deck open and active**.
* In copy mode, the deck has been **saved at least once** (it needs a folder to
  put the copy in) — otherwise the run is refused with `UNSAVED`.
* ScripTree's bundled `lib/combridge/combridge.exe` is present.

---

## 2. The logic (reconstruct-the-tool spec)

A **Strategy-A shim**: a ScripTree form runs `slim_deck.py`; the shim bakes the
form values into a C# Roslyn script rendered from `slim_deck.csx.template`, then
runs it through combridge, which owns the live COM connection to PowerPoint.

* **combridge `run-script` has no argv channel** — a `.csx` only sees the plugin
  globals (`pptApp` / `pptPres` / `pptSlide`) and environment. Form values are
  baked in by replacing `__TOKEN__` placeholders (booleans → `true`/`false`,
  the output-format string via `csharp_literal`).
* **combridge swallows the script's `return` value** — a clean run exits 0. So
  the script prints a first-line **sentinel** and the shim translates it into
  the process exit code.

### Sentinel + status → exit-code contract

First stdout line of the `.csx`:

```
__PPTSLIM__ STATUS=<code> [key=value ...]
```

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Done; report follows (`removed=N mastersRemoved=M before=L after=L-N copy=0/1`). |
| `NODECK` | 2 | No presentation open/active in PowerPoint. |
| `UNSAVED` | 2 | Copy mode requested but the deck has never been saved (no folder for the copy). |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes. |

The shim parses only the `STATUS=` field and maps `NODECK`/`UNSAVED` → exit 2,
everything else → 0. The remaining `key=value` pairs on the sentinel line are
informational and are stripped from the user-visible body.

### The layout-usage keying logic — and why it is conservative

The core problem: **delete only layouts no slide uses, and never delete one a
slide actually uses.** The naive approach — compare each layout object to each
slide's `CustomLayout` — is **unsafe over COM** because RCW (Runtime Callable
Wrapper) object identity is unreliable: two reads of "the same" layout can hand
back different wrappers that compare unequal, so reference comparison can wrongly
classify a used layout as unused and delete it.

Instead we build a **stable, identity-free string key** for every layout:

```
key = <parent master Name>  + "" + <layout Name> + "" + <layout Index>
```

(The code uses an unambiguous separator so e.g. names containing the separator
can't collide.) The Index is the layout's **1-based position within its master**.

1. **Build the used set.** Iterate `work.Slides`; for each slide read
   `slide.CustomLayout`, find its parent master (`cl.Parent as SlideMaster`),
   and add `LayoutKey(masterName, cl.Name, cl.Index)` to a `HashSet`.
2. **Delete candidates.** Iterate each design's
   `SlideMaster.CustomLayouts`; for each layout, derive the **same** key and
   delete it **only if** the key is **not** in the used set.

Why this is safe / conservative:

* If two layouts happen to share a key, both are treated as **used** (kept) —
  a collision can only cause us to under-delete, never over-delete.
* If reading a slide's `CustomLayout` throws, that slide simply doesn't
  contribute a key. Worst case we fail to record a real usage → we keep a
  layout we could have deleted. Again: under-delete, never over-delete.
* **Never delete the last remaining layout of a master** (guard: if
  `master.CustomLayouts.Count <= 1`, stop) — a master with no layouts is
  invalid, and there's no size win worth that risk.
* Every per-layout delete is wrapped in **try/catch** — one stubborn layout
  must not abort the sweep.

This "**on any uncertainty, KEEP**" principle is the whole reason the tool is
safe to run unattended on decks you care about.

### Iterating CustomLayouts BACKWARD

`CustomLayout.Delete()` removes the layout from
`SlideMaster.CustomLayouts` and **reindexes** the collection forward (everything
after the deleted item shifts down one). Iterating **forward** while deleting
would skip the item that slid into the just-deleted slot. So we count down:

```csharp
for (int i = master.CustomLayouts.Count; i >= 1; i--) { ... }
```

(The collection is **1-based**.) The same backward rule applies to
`work.Designs` in the empty-master purge.

### `remove_empty_masters` (opt-in)

After the layout purge, if requested:

* Build the set of master names **used** by any slide (same `cl.Parent` walk).
* Iterate `work.Designs` **backward**; delete a design only when its
  `SlideMaster.CustomLayouts.Count == 0` **and** its master name is **not** in
  the used set.
* **Never delete the last remaining design** (guard:
  `if (work.Designs.Count <= 1) break;`). Wrapped in try/catch.

### OLE object detection — FYI only

Embedded/linked OLE objects (e.g. pasted Excel ranges) are a common, large
source of bloat. The tool **counts** shapes whose
`Type == MsoShapeType.msoEmbeddedOLEObject || msoLinkedOLEObject` across all
slides and **reports the count**, but does **not** touch them. The layout purge
is completely independent of this count — it's surfaced purely so the user knows
where remaining bloat may be hiding.

### What the `.csx` does (step by step)

1. **Guard:** `pptPres is null` → `NODECK`.
2. **Mode:**
   * **Copy mode** (`work_on_copy`): if `pptPres.Path == ""` (never saved) →
     `UNSAVED`. Otherwise `DisplayAlerts = ppAlertsNone`,
     `pptPres.SaveCopyAs(<dir>/<base>_Slimmed.pptx, ppSaveAsOpenXMLPresentation)`,
     then `work = pptApp.Presentations.Open(copyPath, ReadOnly:msoFalse,
     Untitled:msoFalse, WithWindow:msoFalse)` (headless). `copy=1`.
   * **In-place mode:** `work = pptPres`. `copy=0`.
3. **Build the used-layout key set** from `work.Slides` (see above).
4. **Count OLE objects** across `work.Slides` (FYI).
5. **Count layouts before** (`beforeLayouts` = sum of every design's
   `SlideMaster.CustomLayouts.Count`) and, if `remove_unused_layouts`, **delete
   unused layouts backward** with the last-layout guard.
6. **If `remove_empty_masters`,** delete empty/unused designs backward with the
   last-design guard.
7. **Persist:** copy mode → `work.Save()`; in-place → leave UNSAVED (never
   `Save` the original). `finally`: close `work` only if it was the opened copy;
   restore `DisplayAlerts`.
8. **Report:** sentinel
   `__PPTSLIM__ STATUS=OK removed={N} mastersRemoved={M} before={L} after={L-N} copy={0/1}`
   then a markdown/plain-text summary. The summary **explicitly states** that
   image/media recompression is NOT performed.

### Statuses → exit map

`NODECK` / `UNSAVED` → **exit 2**; `OK` (including `removed=0`, a clean no-op) →
**exit 0**.

### PowerPoint COM facts this relies on (see the office-com RAG)

* **`SaveCopyAs` does NOT repoint the active presentation** (contrast Word
  `SaveAs2`). This is what lets copy mode guarantee the open deck is untouched.
* **`Presentations.Open(..., WithWindow:msoFalse)`** opens a deck **headless**
  (no editor window); the returned `Presentation` is the edit target.
* **`Design.SlideMaster.CustomLayouts`** is the per-master, **1-based**
  collection of custom layouts; `CustomLayout.Delete()` reindexes it forward
  (→ iterate backward). `CustomLayout.Index` is its 1-based position;
  `CustomLayout.Parent` is its `SlideMaster`.
* **`Slide.CustomLayout`** is the layout a slide uses; **`work.Designs`** is the
  1-based collection of designs (each wraps one `SlideMaster`); `Design.Delete()`
  removes a design.
* **RCW identity is unreliable** → key layouts by (master Name + layout Name +
  Index), never by reference equality.
* `MsoTriState` / `MsoShapeType` (incl. `msoEmbeddedOLEObject`,
  `msoLinkedOLEObject`) come from `Microsoft.Office.Core`; the `Pp*` enums
  (`PpSaveAsFileType`, `PpAlertLevel`) from
  `Microsoft.Office.Interop.PowerPoint`. Both are in the PowerPoint plugin's
  default usings.
* Modal-dialog hang: `SaveCopyAs` / `Open` / `Save` can raise modal prompts on a
  headless instance — hence `DisplayAlerts = ppAlertsNone` around them, restored
  in `finally`.

---

## 3. Files in this app

| File | Role |
|---|---|
| `deck-slimmer.scriptree` | The form (4 params; a slide-with-compress-arrows icon embedded as PNG). |
| `deck-slimmer.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config. |
| `slim_deck.py` | The Strategy-A shim. |
| `slim_deck.csx.template` | The Roslyn template with `__TOKEN__` placeholders. |
| `examples/` | `make_example.py` + `sample_bloated.pptx` (11 layouts, 2 used) + README. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
slim_deck.py
  [--remove-unused-layouts]   (flag; default ON in the form)
  [--remove-empty-masters]    (flag; default OFF — opt-in)
  [--work-on-copy]            (flag; default ON in the form)
  --output-format markdown|text
```

The three booleans use the conditional `{id?--flag}` token form (the flag is
emitted only when ticked); `--output-format` passes through a
`["--flag","{id}"]` token group. `store_true` argparse on the shim side.

---

## 4. Editing / maintenance notes

* **Never claim image/media compression.** It is technically not possible
  fire-and-forget over COM (interactive Compress Pictures dialog only). Keep the
  description, README, and report wording honest about layout/master removal
  being the sole mechanism.
* **Conservative deletion is load-bearing.** Keep the "key by master+name+index,
  keep on any uncertainty, never delete the last layout/design" rules. If you
  ever make deletion more aggressive, gate it behind a clearly-labelled opt-in
  and document the new risk.
* **Iterate CustomLayouts / Designs BACKWARD** when deleting — forward iteration
  skips items after a `Delete()` reindex.
* **Keep the per-layout / per-slide try/catch guards** — a single stubborn
  layout or system slide must not abort the sweep.
* **Copy mode vs in-place:** only the opened copy is ever `Save()`d or
  `Close()`d. The in-place path deliberately leaves the open deck UNSAVED so the
  user reviews before committing.
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree validate <path>`.
* **Offline render-check:** render the template with sample values and grep for
  `__[A-Z_]+__` — only the `__PPTSLIM__` sentinel should match; anything else is
  an unfilled placeholder.
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path.
