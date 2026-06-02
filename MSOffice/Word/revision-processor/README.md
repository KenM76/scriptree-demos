# Tracked-Changes Processing Master (Word)

Force-accept **or** force-reject **all** tracked changes (revisions) in the
**open Word document** in one pass, clearing the document's revision history so
it is clean to send out. Useful for finalising a reviewed document before it
leaves the building.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

## The relationship to the Style Sanitizer (read this)

This tool is the **inverse** of the Corporate Style Sanitizer
(`../style-sanitizer/`). The sanitizer **refuses to run** on a document that has
tracked changes — bulk-editing a document under review would bury it in
revision marks. This tool **specifically targets** documents that *have*
revisions and clears them. They are complementary: clear the revisions here
first, then sanitise.

---

## 1. What the user sees (end-user guide)

### The form fields

| Field | Meaning |
|---|---|
| **Action** (radio; default *Accept all changes*) | **Accept all changes** incorporates every insertion / deletion / format change into the text as if approved (the same as Review → Accept All in Word). **Reject all changes** reverts every change, restoring the document to its pre-revision state (Review → Reject All). Either way, **all** revisions are cleared in one pass. |
| **Work on a copy** (default ON) | SAFETY GUARD. When on, the document is first saved as `<name>_Revisions_Processed.docx` in the same folder and the accept/reject lands on that copy — the original is never modified. When off, the result is applied to the open document **in memory and left unsaved** for you to review (Ctrl+Z still undoes it). |
| **Output format** | `Markdown` or `Plain text` for the result summary. |

### What it does to your file

* **Work on a copy = ON (default):** a `<name>_Revisions_Processed.docx` copy is
  created next to the original; the accept/reject happens there; the copy is
  saved. Your original file is never touched.
* **Work on a copy = OFF:** the accept/reject is applied to the open document
  **in memory and left UNSAVED** — review in Word and save (or undo) yourself.
  The tool never silently overwrites the original on disk.
* After processing, **Track Changes is turned OFF** on the processed document so
  it is not left in tracking mode (otherwise the next edit would immediately
  start marking up the freshly-cleaned document again).
* If the document has **no tracked changes**, the tool reports success and does
  **nothing** — no copy is made and the document is not modified. A document
  already free of revisions is the desired end state, not an error.

### Prerequisites

* The target document is **already open** in a running Word instance.
* For "Work on a copy", the document has been **saved at least once** (it needs a
  folder for the copy) — otherwise the run is refused.
* ScripTree's bundled `lib/combridge/combridge.exe` is present.

---

## 2. The logic (reconstruct-the-tool spec)

A **Strategy-A shim**: a ScripTree form runs `process_revisions.py`; the shim
bakes the form values into a C# Roslyn script rendered from
`process_revisions.csx.template`, then runs it through combridge, which owns the
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
__WORDREV__ STATUS=<code> [key=value ...]
```

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` (with `revisions=N`, N>0) | 0 | Revisions processed; report follows (`revisions=N action=accept|reject copy=0/1`). |
| `OK` (with `revisions=0`) | 0 | **Document had no tracked changes — nothing to do.** This is a *success*, not a precondition failure: it is reported under `STATUS=OK` and maps to exit 0. No copy is made; the document is not modified. |
| `NODOC` | 2 | No document open in Word. |
| `UNSAVED` | 2 | "Work on a copy" requested but the document was never saved. |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes. |

> The shim's `PRECONDITION_FAILS` tuple is `("NODOC", "UNSAVED")` only —
> `revisions=0` is deliberately **not** in it, so it returns 0.

### What the `.csx` does, in order

1. **No-document guard:** `wdDoc` null → `__WORDREV__ STATUS=NODOC`, return 0
   (the shim maps NODOC → exit 2).
2. **Count revisions (guarded):** `int revs = 0; try { revs =
   wdDoc.Revisions.Count; } catch { }`. If `revs == 0` → `__WORDREV__ STATUS=OK
   revisions=0 action=… copy=…` plus a "no tracked changes — nothing to do"
   summary, return 0. **No copy is made and the document is not touched.**
3. **Copy guard (if "work on a copy"):** if `wdDoc.Path == ""` (never saved) →
   `__WORDREV__ STATUS=UNSAVED`, return 0 (shim → exit 2). Otherwise build
   `copyPath = <dir>/<baseName>_Revisions_Processed.docx` and, with
   `DisplayAlerts` off, `wdDoc.SaveAs2(copyPath, wdFormatXMLDocument)`.
   **`SaveAs2` repoints the active document at the copy**, so the accept/reject
   below and the final `Save()` land on the copy — the original file is
   untouched.
4. **Apply (in a try/finally that restores DisplayAlerts):**
   * `action == "reject"` → `wdDoc.RejectAllRevisions();`
   * else → `wdDoc.AcceptAllRevisions();`
   * then `wdDoc.TrackRevisions = false;` so the processed document isn't left
     in tracking mode.
   * in copy-mode only, `wdDoc.Save();`.
5. **Report:** sentinel `__WORDREV__ STATUS=OK revisions={N} action={accept|reject}
   copy={0/1}` then the markdown/plain-text summary (document name, count
   processed, action, that tracking was turned off, and the copy path / in-place
   note).

### Why `AcceptAllRevisions()` / `RejectAllRevisions()` (not a loop)

The **Document-level** `AcceptAllRevisions()` / `RejectAllRevisions()` methods
process every revision in the document body in one atomic call — exactly Review →
Accept All / Reject All in the ribbon. Looping `wdDoc.Revisions` and calling
`Accept()`/`Reject()` per item is slower and fragile (the collection mutates as
you accept/reject). The collection-level `wdDoc.Revisions.AcceptAll()` /
`.RejectAll()` exist and are equivalent; this tool uses the Document-level form.

> **NOTE (flagged for verification):** these method names
> (`AcceptAllRevisions`, `RejectAllRevisions`) and `TrackRevisions` were
> confirmed present in the bundled `Microsoft.Office.Interop.Word.dll`. They are
> the standard Word object-model members, but the live COM round-trip is pending
> verification against a running Word like every other app in this catalog.

### Word COM facts this relies on (see the office-com RAG)

* **`wdDoc.Revisions.Count`** is the revision count; wrap the read in try/catch
  (it can throw on some protected/odd documents).
* **`AcceptAllRevisions()` / `RejectAllRevisions()`** clear all revisions in one
  call.
* **`TrackRevisions = false`** stops the document tracking further edits — set it
  *after* the accept/reject so the processed document is left clean, not armed.
* **`SaveAs2(path, wdFormatXMLDocument)` repoints `wdDoc`** at the new file — the
  basis of the "work on a copy" guard.
* **`wdDoc.Path` is `""` until the document has been saved** — the never-saved
  copy-mode refusal.
* **Save paths raise modal dialogs that hang a hidden COM-launched Word** — set
  `wdApp.DisplayAlerts = wdAlertsNone` around `SaveAs2`/`Save`, restored in a
  `finally`.

---

## 3. Files in this app

| File | Role |
|---|---|
| `revision-processor.scriptree` | The form (3 params; a document-page-with-green-check-and-red-X PNG icon embedded — the accept/reject duality). Built by a throwaway `_build_scriptree.py` (since deleted). |
| `revision-processor.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config (popups on; command-preview off). |
| `process_revisions.py` | The Strategy-A shim. |
| `process_revisions.csx.template` | The Roslyn template with `__TOKEN__` placeholders. |
| `examples/` | A generator + a `sample_review.docx` with real tracked changes + a README. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
process_revisions.py
  --action accept|reject          (radio; default accept)
  [--work-on-copy]                (flag; default ON in the form)
  --output-format markdown|text
```

`--action` and `--output-format` use `["--flag","{id}"]` token groups; the
boolean uses the `{id?--flag}` conditional form (`store_true` in argparse).

---

## 4. Editing / maintenance notes

* **The "work on a copy" guard is the headline safety property** — don't let an
  edit make the in-place path overwrite the original on disk. In-place mode must
  leave the document unsaved.
* **`revisions=0` is OK, not a failure** — keep it under `STATUS=OK` (exit 0).
  Putting it in the shim's `PRECONDITION_FAILS` tuple would wrongly make
  "document already clean" report as an error.
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree validate <path>`. Expect "widgets: checkbox, dropdown,
  radio".
* **Offline render-check:** render the template with sample values (reject +
  in-place + text, and accept + copy + markdown) and grep for `__[A-Z_]+__` —
  only the `__WORDREV__` sentinel should match; anything else is an unfilled
  placeholder.
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path. Run from the
  source repo (no combridge bundle) and the shim correctly exits 1 with
  "could not locate ...".
