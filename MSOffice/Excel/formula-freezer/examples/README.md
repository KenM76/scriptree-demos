# Example — Formula-to-Value Freezer

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates the sample. |
| `sample_formulas.xlsx` | A two-sheet workbook full of real formulas (see below). |

### What's in the sample, and why

**Sheet `Sales`** — a small table with a header row `Item | Qty | Price | Total`
and 4 data rows. The **Total** column is a per-row formula `=B*C` (Qty × Price).
Underneath is a summary block:

| Label | Cell | Formula |
|---|---|---|
| Units | `Sales!B7` | `=SUM(B2:B5)` |
| Avg price | `Sales!B8` | `=AVERAGE(C2:C5)` |
| Revenue | `Sales!B9` | `=SUM(D2:D5)` (a SUM over the formula column) |

**Sheet `Summary`** — pulls figures from `Sales` via **cross-sheet references**:

| Label | Cell | Formula |
|---|---|---|
| Total revenue | `Summary!B2` | `=Sales!B9` |
| Total units | `Summary!B3` | `=Sales!B7` |
| Average price | `Summary!B4` | `=Sales!B8` |
| Revenue / unit | `Summary!B5` | `=B2/B3` (a formula built on other formulas) |

That mix — per-row arithmetic, `SUM`/`AVERAGE` aggregates, cross-sheet links, and
a formula that depends on other formulas — is exactly what "freezing" is for:
afterwards none of those cells recalculates, because each holds the constant it
last produced.

## How to try it

1. Open `sample_formulas.xlsx` in Excel **and save it once** (it's freshly
   generated; copy mode needs a saved file to put the frozen copy beside). The
   **Sales** sheet is active.
2. Run **Formula-to-Value Freezer** from ScripTree.

Form values (the safe defaults):

| Field | Value |
|---|---|
| Scope | `All worksheets` |
| Work on a copy | *(checked)* |
| Output format | `markdown` |

**Expected result** — a sibling `sample_formulas_Frozen.xlsx` on disk in which:

* `Sales!D2:D5` are now the **constants** `540`, `960`, `435`, `899.55` — no
  longer `=B*C`;
* `Sales!B7` is `305` (the frozen `SUM`), `Sales!B8` is `10.935` (the frozen
  `AVERAGE`), `Sales!B9` is `2834.55` (the frozen revenue `SUM`);
* every `Summary` cell is a constant too — `=Sales!B9` became the value, and
  `Revenue / unit` became `2834.55 / 305 = 9.2936…`;
* the **number formats are unchanged** (the price/total cells still display as
  numbers/currency exactly as before);
* your original `sample_formulas.xlsx` is **untouched** — open it and the
  formulas are all still there.

You can confirm the freeze by reopening the `_Frozen` copy: every cell that was
a formula now shows a literal value in the formula bar.

### Variations to try

* **Active sheet only:** make **Summary** the active tab, set **Scope** =
  `Active sheet only`, and run. Only `Summary` freezes; `Sales` keeps its
  formulas. (Note the frozen `Summary` cells now hold the values they had at
  freeze time — they no longer track later edits to `Sales`.)
* **In place:** untick **Work on a copy**. The open workbook is frozen in
  memory and left **UNSAVED** — review it, then save (or press Ctrl+Z to undo)
  yourself. Your file on disk is not changed until you save.
* **The protection refusal:** protect a sheet (Review → Protect Sheet) and run
  with that sheet in scope. The tool **refuses** with a `PROTECTED` message
  naming the sheet, before making any change — a protected sheet's cells can't
  be overwritten with values.
* **The unsaved refusal:** on a brand-new never-saved workbook, leave **Work on
  a copy** on and run — the tool refuses with `UNSAVED` (there's no folder to
  put the copy in). Save once, or untick the copy option.

## What this demonstrates

* Replacing every formula with its current value while preserving number
  formats — across all sheets or just the active one.
* The copy-mode safety default: the deliverable is a saved `_Frozen.xlsx`; your
  original stays open and untouched.
* The protected-sheet and never-saved refusals.

> Like every app in this catalog, the freezer is pending live verification
> against a real Excel. The scenario above describes the exact behaviour to
> expect.
