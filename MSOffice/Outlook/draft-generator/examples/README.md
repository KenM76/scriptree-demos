# Example — Draft Generator from Template (Outlook)

This folder gives you a ready-to-run mailing list and the exact form values to
try the tool against it.

> **This example CREATES ITEMS in your live mailbox — but only DRAFTS.** Running
> it produces three **unsent draft emails** in your Outlook Drafts folder.
> Nothing is sent. The tool never calls Send; it only ever calls `Save()`, which
> lands an unsent draft. Review (and delete, if you like) the three drafts
> afterward.

## The sample file

`sample_list.xlsx` (regenerate any time with `python make_example.py`) has a
header row and **three** data rows on its active sheet `Recipients`:

| Email | Name | Company |
|---|---|---|
| ada@example.com | Ada Lovelace | Analytical Engines Ltd |
| alan@example.com | Alan Turing | Bletchley Works |
| grace@example.com | Grace Hopper | Compiler Co |

The addresses are obviously fake (`example.com`). Because the tool only creates
**drafts**, no mail goes anywhere even though the addresses aren't real.

## How to try it

1. Make sure **Outlook is running** (a profile loaded — no window need be open).
2. Run **Draft Generator from Template** from ScripTree.

| Field | Value | Effect |
|---|---|---|
| Mailing list (.xlsx) | *(browse to `sample_list.xlsx`)* | the three-row list above |
| Recipient column header | `Email` | the column holding addresses (case-insensitive match) |
| Subject template | `Statement for {Name} at {Company}` | `{Name}`/`{Company}` are replaced per row |
| Body template | *(the multi-line block below)* | `{Name}` replaced per row |
| Maximum drafts (cap 50) | `50` | all 3 rows fit under the cap |
| Output format | `markdown` | rendered report (`text` for plain monospace) |

A multi-line **body template** to paste in:

```
Dear {Name},

Your latest account statement is ready for review. Please find the
summary attached, and reply to this message with any questions.

Kind regards,
Accounts Team
```

## Expected result

Three **drafts** appear in your Outlook **Drafts** folder, one per row, with the
`{Name}` / `{Company}` tokens filled in:

| To | Subject | Body opens with |
|---|---|---|
| ada@example.com | Statement for Ada Lovelace at Analytical Engines Ltd | `Dear Ada Lovelace,` |
| alan@example.com | Statement for Alan Turing at Bletchley Works | `Dear Alan Turing,` |
| grace@example.com | Statement for Grace Hopper at Compiler Co | `Dear Grace Hopper,` |

Nothing is sent.

## Expected output (markdown mode)

```
# Draft generation

- **Mailing list:** `...\examples\sample_list.xlsx`
- **Recipients processed:** 3
- **Drafts created:** 3 (in the 'Drafts' folder)

| # | To | Subject |
|--:|---|---|
| 1 | ada@example.com | Statement for Ada Lovelace at Analytical Engines Ltd |
| 2 | alan@example.com | Statement for Alan Turing at Bletchley Works |
| 3 | grace@example.com | Statement for Grace Hopper at Compiler Co |

> **No emails were sent.** Each item above is a DRAFT — review them in your
> Outlook Drafts folder and send them yourself when you're ready.
```

## Variations to try

* **Skip a row:** clear the `Email` cell on one row in `sample_list.xlsx` — that
  row is skipped (blank recipient), and you get 2 drafts instead of 3.
* **A token with no column:** put `{Region}` in the body. There is no `Region`
  header, so `{Region}` is left **as-is** in the draft (unknown braces aren't
  touched).
* **Lower the cap:** set **Maximum drafts** to `2` — only the first 2 usable rows
  are drafted. (The cap can never exceed 50, the absolute maximum.)
* **Wrong column name:** set **Recipient column header** to `Mail` — there's no
  such header, so the run is refused with a clear error (exit 2) and no drafts.

## What this demonstrates

* A mail merge that **stops at drafts** — personalised per-row drafts you review
  before sending.
* `{Column}` token rendering from the spreadsheet headers (case-sensitive on the
  exact header text; the recipient-column lookup is case-insensitive).
* Blank-recipient rows being skipped.
* The **drafts-only safety contract** and the **50-draft hard cap**: the tool
  only ever calls `Save()` (never `Send()`), and never creates more than 50
  drafts in a run.

> Pending live verification against a real Outlook profile; the scenario above
> describes the exact report shape and Drafts-folder result to expect.
