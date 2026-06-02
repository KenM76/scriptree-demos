# Example — Batch Find & Replace

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates the sample. |
| `sample_memo.docx` | A status memo seeded with the word **Foobar** in several forms so each matching option behaves differently. |

The sample contains exactly: **5** standalone `Foobar`, **1** lower-case
`foobar`, **1** `Foobars`, **1** `Foobaristas`.

## How to try it

1. Open `sample_memo.docx` in Word.
2. Run **Batch Find & Replace** from ScripTree with **Find what** =
   `Foobar` and **Replace with** = `Widget`.

By default **Work on a copy** is on, so edits land in a new
`sample_memo_Replacements.docx` next to the sample and your open document
is never touched.

### The count changes with the options

| Match case | Whole words only | Replacements | Why |
|---|---|---|---|
| off | off | **8** | every substring "foobar" (any case): the 5 + `foobar` + `Foobars` + `Foobaristas` |
| **on** | off | **7** | drops the lower-case `foobar`; `Foobars`/`Foobaristas` still match (they contain "Foobar") |
| off | **on** | **6** | drops `Foobars`/`Foobaristas` (not whole words); keeps the 5 + `foobar` |
| **on** | **on** | **5** | only the standalone, correctly-cased `Foobar` tokens |

So the precise "replace whole-word, case-sensitive Foobar → Widget" run
reports **5 replacements**.

### Delete instead of replace

Leave **Replace with** blank to *delete* every match instead of
substituting — the report still gives the exact count removed.

### Edit in place (no copy)

Untick **Work on a copy** to edit the open document directly. The edits
are left **unsaved** in Word so you can review (or Ctrl+Z) before saving
— the tool will not overwrite your original on disk.

## What this demonstrates

* An exact, auditable replacement count (Word's dialog doesn't give one).
* How Match case and Whole words only each change what matches.
* The safe-by-default "work on a copy" behaviour.
