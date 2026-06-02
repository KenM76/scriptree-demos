# Batch Find & Replace (Word)

Find every occurrence of a phrase in the **open Word document** and replace it in
one pass, then report exactly how many replacements were made. Faster and more
auditable than Word's built-in dialog when you want a clean count and a safe copy.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

---

## 1. What the user sees (end-user guide)

### The form fields

The form groups its fields into three collapsible sections — **Find & replace**, **Matching**, and **Output & safety**.

| Section | Field | Meaning |
|---|---|---|
| Find & replace | **Find what** (required) | The text to search for. With "Use wildcards" on, this is a Word wildcard pattern rather than a literal phrase. |
| Find & replace | **Replace with** | The replacement text. Leave blank to **delete** every occurrence of the Find text. |
| Matching | **Match case** | When on, `Report` matches `Report` but not `report`. |
| Matching | **Whole words only** | When on, searching `cat` does not match the `cat` inside `category`. |
| Matching | **Use wildcards** | Treat "Find what" as a Word wildcard pattern (the same engine as Word's "Use wildcards" checkbox). |
| Output & safety | **Work on a copy** (default ON) | SAFETY GUARD. When on, the document is first saved as `<name>_Replacements.docx` in the same folder and all edits land on that copy — the original is never modified. When off, edits are made to the open document in memory and **left unsaved** for you to review and save. |
| Output & safety | **Output format** | `Markdown` or `Plain text` for the result summary. |

### What it does to your file

* **Work on a copy = ON (default):** a `<name>_Replacements.docx` copy is created
  next to the original; all replacements are made there; the copy is saved. Your
  original file is never touched.
* **Work on a copy = OFF:** replacements are made in the open document **in
  memory and left UNSAVED** — review in Word and save (or undo) yourself. The
  tool never silently overwrites the original on disk.

### Prerequisites

* The target document is **already open** in a running Word instance.
* For "Work on a copy", the document has been **saved at least once** (it needs a
  folder to put the copy in) — otherwise the run is refused.
* ScripTree's bundled `lib/combridge/combridge.exe` is present.

---

## 2. The logic (reconstruct-the-tool spec)

A **Strategy-A shim**: a ScripTree form runs `word_find_replace.py`; the shim
bakes the form values into a C# Roslyn script rendered from
`word_find_replace.csx.template`, then runs it through combridge, which owns the
live COM connection to Word.

* **combridge `run-script` has no argv channel** — a `.csx` only sees the plugin
  globals (`wdApp` / `wdDoc`) and environment. Form values are baked in by
  replacing `__TOKEN__` placeholders.
* **combridge swallows the script's `return` value** — clean run exits 0 (compile=3
  / throw=4 / host=5 differ). So the script prints a first-line **sentinel** and
  the shim translates it into the process exit code.

### Sentinel + status → exit-code contract

First stdout line of the `.csx`:

```
__WORDREP__ STATUS=<code> [key=value ...]
```

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Replacements ran; report follows (`replaced=N copy=0/1`). |
| `NODOC` | 2 | No document open in Word. |
| `EMPTY_FIND` | 2 | The Find field was empty. |
| `UNSAVED` | 2 | "Work on a copy" requested but the document was never saved. |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes. |

### What the `.csx` does

1. **Guards:** `wdDoc` null → `NODOC`; empty find text → `EMPTY_FIND`.
2. **Copy guard (if "work on a copy"):** if `wdDoc.Path == ""` (never saved) →
   `UNSAVED`. Otherwise build `copyPath = <dir>/<baseName>_Replacements.docx` and,
   with `DisplayAlerts` off, `wdDoc.SaveAs2(copyPath, wdFormatXMLDocument)`.
   **`SaveAs2` repoints the active document at the copy**, so every edit below and
   the final `Save()` land on the copy — the original file is untouched.
3. **The replace loop (exact count, infinite-loop-proof):** `Range rng =
   wdDoc.Content`; loop: configure `rng.Find` (`ClearFormatting`, `Text=findText`,
   `Forward=true`, `Wrap=wdFindStop` so it doesn't wrap, plus the match-case /
   whole-word / wildcards flags); `if (!find.Execute()) break;` else set
   `rng.Text = replaceText`, **collapse PAST** the replacement
   (`rng.Collapse(wdCollapseEnd)` then `rng.End = wdDoc.Content.End`), and
   `replaced++`. Collapsing past the replacement is what stops a replacement that
   *contains* the search text from looping forever. A `SAFETY_CAP` of 1,000,000
   is a final backstop.
4. **Persist only in copy-mode:** `wdDoc.Save()` (DisplayAlerts off). In-place
   mode deliberately leaves the edits unsaved.
5. **Report:** sentinel `__WORDREP__ STATUS=OK replaced={N} copy={0/1}` then the
   markdown/plain-text summary (find, replace, count, options, and whether it
   worked on a copy + the copy path / original name).

### Word COM facts this relies on (see the office-com RAG)

* `Find.Execute` returns a **bool, not a count** — hence the find-one-then-collapse
  loop.
* `SaveAs2(path, wdFormatXMLDocument)` **repoints `wdDoc`** at the new file — the
  basis of the "work on a copy" guard.
* `wdDoc.Path` is `""` until the document has been saved.
* `Worksheet`/document save paths raise **modal** dialogs that hang a hidden
  COM-launched Word — set `wdApp.DisplayAlerts = wdAlertsNone` around them.

---

## 3. Files in this app

| File | Role |
|---|---|
| `batch-find-replace.scriptree` | The form (7 params; `edit`/pencil icon embedded as PNG — `edit` not `search`, because this MUTATES the doc). |
| `batch-find-replace.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config. |
| `word_find_replace.py` | The Strategy-A shim. |
| `word_find_replace.csx.template` | The Roslyn template with `__TOKEN__` placeholders. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
word_find_replace.py
  --find <str>                  (required)
  --replace <str>               (blank = delete matches)
  [--match-case]                (flag)
  [--whole-word]                (flag)
  [--use-wildcards]             (flag)
  [--work-on-copy]              (flag; default ON in the form)
  --output-format markdown|text
```

`--find` / `--replace` use `["--flag","{id}"]` token groups with `no_split: true`
on the string params so phrases with spaces stay single tokens; the booleans use
the `{id?--flag}` conditional form.

---

## 4. Editing / maintenance notes

* **The "work on a copy" guard is the headline safety property** — don't let an
  edit make the in-place path overwrite the original on disk. In-place mode must
  leave the document unsaved.
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree.cli.validate <path>`.
* **Offline render-check:** render the template with sample values (quotes,
  backslashes, tabs in the find text) and grep for `__[A-Z_]+__` — only the
  `__WORDREP__` sentinel should match; anything else is an unfilled placeholder.
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path.
