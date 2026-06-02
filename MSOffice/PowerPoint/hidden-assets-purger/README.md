# Hidden Assets & Notes Purger (PowerPoint)

Produce a **sanitized copy** of the open PowerPoint deck with speaker notes,
reviewer comments, author/document metadata, and (optionally) hidden slides
stripped out. Use it before sending a deck outside your organisation so private
talking points, internal review threads, and the author's identity don't leak.
The deck you have open is **never modified** — the tool writes a sibling
`<name>_Sanitized.pptx` and does all stripping on that copy.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

This is the first **PowerPoint** app in the catalog — it stands up the
`powerpoint` combridge plugin pattern (globals `pptApp` / `pptPres` / `pptSlide`).

---

## 1. What the user sees (end-user guide)

### The form fields

The form groups its fields into two collapsible sections — **What to strip** and **Slides & output**.

| Section | Field | Default | Meaning |
|---|---|---|---|
| What to strip | **Remove speaker notes** | ON | Clear the speaker-notes text from every slide in the copy. Notes often hold private talking points, internal figures, or reminders. |
| What to strip | **Remove comments** | ON | Delete reviewer comments from every slide in the copy. Reaches PowerPoint's **classic** comment collection; very recent **modern threaded** comments may not all be exposed to automation and could survive — spot-check if your deck uses them heavily. |
| What to strip | **Strip author & document metadata** | ON | Remove author, company, last-saved-by, revision history and other document properties from the copy (PowerPoint's "Remove All Document Information"). |
| Slides & output | **Permanently delete hidden slides** | OFF (opt-in) | Delete every slide marked hidden from the copy. Hidden slides frequently hold backup numbers, draft content, or appendix material. Off by default because deletion on the copy is permanent (your original is still untouched either way). |
| Slides & output | **Output format** | `Markdown` | `Markdown` (paste into a ticket) or `Plain text` (aligned list). |

### What it does to your deck

**Nothing to the original.** The open deck is strictly untouched. The tool
saves a copy named `<name>_Sanitized.pptx` in the **same folder** as the
original and performs every strip operation on that copy, then leaves the copy
saved and closed. You open the copy yourself to review/send it.

### Prerequisites

* PowerPoint is **running with the target deck open and active**.
* The deck has been **saved at least once** (it needs a folder to put the copy
  in) — otherwise the run is refused with `UNSAVED`.
* ScripTree's bundled `lib/combridge/combridge.exe` is present.

---

## 2. The logic (reconstruct-the-tool spec)

A **Strategy-A shim**: a ScripTree form runs `ppt_purger.py`; the shim bakes the
form values into a C# Roslyn script rendered from `ppt_purger.csx.template`, then
runs it through combridge, which owns the live COM connection to PowerPoint.

* **combridge `run-script` has no argv channel** — a `.csx` only sees the plugin
  globals (`pptApp` / `pptPres` / `pptSlide`) and environment. Form values are
  baked in by replacing `__TOKEN__` placeholders.
* **combridge swallows the script's `return` value** — clean run exits 0. So the
  script prints a first-line **sentinel** and the shim translates it into the
  process exit code.

### Sentinel + status → exit-code contract

First stdout line of the `.csx`:

```
__PPTPURGE__ STATUS=<code> [key=value ...]
```

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Sanitized copy written; report follows (`notes=N comments=C hidden=H meta=0/1`). |
| `NODECK` | 2 | No presentation open/active in PowerPoint. |
| `UNSAVED` | 2 | Deck has never been saved, so there is no folder for the copy. |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes. |

### The headline safety property — the open deck is never mutated

This is the whole value proposition and it is **load-bearing**: the file the
user is looking at must be provably untouched.

* **`Presentation.SaveCopyAs(path, fileFormat)` writes a copy WITHOUT repointing
  the active presentation.** This is the critical difference from Word's
  `SaveAs2` (which *does* repoint `wdDoc` at the new file). After `SaveCopyAs`,
  `pptPres` still refers to the original, still open, still unmodified.
* We then **open the copy headless** with
  `pptApp.Presentations.Open(copyPath, ReadOnly:msoFalse, Untitled:msoFalse, WithWindow:msoFalse)`
  and do **all** stripping on that returned `Presentation`, then `Save()` and
  `Close()` it. The original `pptPres` is never an edit target.
* Therefore there is **no in-place mode** and **no destination picker** — the
  shim/`.csx` derive the copy path from the active presentation's name. Simpler
  and safer than the Word tool's copy-toggle.

### What the `.csx` does

1. **Guard:** `pptPres is null` → `NODECK`.
2. **Guard:** `pptPres.Path == ""` (never saved) → `UNSAVED`.
3. **Derive copy path:** `<dir>/<baseNameWithoutExt>_Sanitized.pptx` where
   `dir = pptPres.Path`, `baseName = GetFileNameWithoutExtension(pptPres.Name)`.
4. **DisplayAlerts off** (`pptApp.DisplayAlerts = PpAlertLevel.ppAlertsNone`,
   saved/restored in a `finally`) — a headless COM instance hangs on any modal
   prompt (overwrite confirmation, etc.).
5. **`SaveCopyAs`** the copy (`PpSaveAsFileType.ppSaveAsOpenXMLPresentation`),
   then **`Presentations.Open(... WithWindow:msoFalse)`** the copy as `work`.
6. **Strip operations on `work`** (order matters — hidden slides first):
   * **(a) Hidden slides** (if opted in): iterate **BACKWARD**
     `for (int i = work.Slides.Count; i >= 1; i--)`; if
     `slide.SlideShowTransition.Hidden == MsoTriState.msoTrue` → `slide.Delete()`.
     Backward iteration means deleting slide `i` never shifts the index of a
     slide still to be visited. Deleting first means notes/comments on
     soon-to-be-deleted slides aren't counted as "cleared".
   * **(b) Speaker notes** (if on): per surviving slide, iterate
     `slide.NotesPage.Shapes`; for each shape that is
     `Type == MsoShapeType.msoPlaceholder` **and**
     `PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody` **and** has
     text, set `TextFrame.TextRange.Text = ""`. The body placeholder is the only
     one holding notes text (the other placeholder is the slide-image thumbnail);
     checking `msoPlaceholder` first avoids the throw from reading
     `PlaceholderFormat` on a non-placeholder shape.
   * **(c) Comments** (if on): per slide,
     `while (slide.Comments.Count > 0) slide.Comments[1].Delete();` —
     `Comments` is **1-based**; always delete `[1]` and let the collection
     shrink. Reaches the **classic** comment collection; **modern threaded
     comments may not be fully exposed** to this API.
   * **(d) Metadata** (if on): `work.RemoveDocumentInformation(PpRemoveDocInfoType.ppRDIAll)`.
   * Every per-slide / per-enumeration step is wrapped in **try/catch** so one
     stubborn system/search slide that throws on `NotesPage`/`Comments` doesn't
     abort the whole sweep.
7. **Persist + close:** `work.Save()` then (in `finally`) `work.Close()`.
8. **Report:** sentinel
   `__PPTPURGE__ STATUS=OK notes={N} comments={C} hidden={H} meta={0/1}` then a
   markdown/plain-text summary naming the untouched source, the copy path, and
   per-option counts (each option shows `(skipped)` when its checkbox was off).

### PowerPoint COM facts this relies on (see the office-com RAG)

* **`SaveCopyAs` does NOT repoint the active presentation** (contrast Word
  `SaveAs2`). This is what lets us guarantee the open deck is untouched.
* **`Presentations.Open(..., WithWindow:msoFalse)`** opens a deck **headless**
  (no editor window) for background processing; the returned `Presentation` is
  the edit target.
* **`Slide.Comments` is 1-based**; delete `[1]` in a `while Count>0` loop.
  Classic vs modern threaded comments: modern comments may not surface here.
* **`NotesPage.Shapes`** holds a slide-image placeholder and a body placeholder;
  only `ppPlaceholderBody` carries notes text. Reading `PlaceholderFormat` on a
  non-placeholder shape **throws** — gate on `MsoShapeType.msoPlaceholder` first.
* **Hidden flag** is `slide.SlideShowTransition.Hidden == MsoTriState.msoTrue`.
* `MsoTriState` / `MsoShapeType` come from `Microsoft.Office.Core`, which is in
  the PowerPoint plugin's default usings alongside
  `Microsoft.Office.Interop.PowerPoint`.
* Modal-dialog hang: even though this is mostly a copy-and-process path, the
  `SaveCopyAs` / `Open` / `Save` calls can raise modal prompts on a headless
  instance — hence `DisplayAlerts = ppAlertsNone` around them.

---

## 3. Files in this app

| File | Role |
|---|---|
| `hidden-assets-purger.scriptree` | The form (5 params; `shield` icon embedded as PNG — `shield` = "make X safe", the sanitizer/leak-prevention archetype). |
| `hidden-assets-purger.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config. |
| `ppt_purger.py` | The Strategy-A shim. |
| `ppt_purger.csx.template` | The Roslyn template with `__TOKEN__` placeholders. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
ppt_purger.py
  [--strip-notes]            (flag; default ON in the form)
  [--strip-comments]         (flag; default ON in the form)
  [--strip-metadata]         (flag; default ON in the form)
  [--delete-hidden-slides]   (flag; default OFF — opt-in)
  --output-format markdown|text
```

The four booleans use the conditional `{id?--flag}` token form (the flag is
emitted only when ticked); `--output-format` passes through a
`["--flag","{id}"]` token group.

---

## 4. Editing / maintenance notes

* **Never add an in-place mode.** The "we only ever touch a copy" guarantee is
  the entire reason this tool is safe to run on a deck you care about. If a
  future variant must edit in place, that re-opens the modal-hang question and
  loses the headline property — make it a separate tool instead.
* **Strip order is load-bearing:** delete hidden slides FIRST (backward),
  *then* clear notes/comments on the survivors, so counts are honest and index
  shifts never skip a slide.
* **Keep the per-slide try/catch guards** — system/search/odd slides throw on
  `NotesPage` or `Comments`; one throw must not abort the sweep.
* **Modern threaded comments** are a known limitation (classic `Slide.Comments`
  may not expose them). If combridge/PowerPoint later surfaces them, extend
  step (c); until then the form copy says so.
* **Edge case — all slides hidden + delete-hidden on:** the copy can end up with
  zero slides. PowerPoint tolerates an empty deck via COM; it is an honest
  result of the user's opt-in, so it is not specially guarded.
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree validate <path>`.
* **Offline render-check:** render the template with sample values and grep for
  `__[A-Z_]+__` — only the `__PPTPURGE__` sentinel should match; anything else
  is an unfilled placeholder.
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path.
