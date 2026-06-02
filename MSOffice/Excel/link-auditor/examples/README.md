# Example — Broken External Link Auditor

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates both workbooks. **Run it from the folder you'll actually open the files from** — the link target is an absolute path baked in at generation time (see below). |
| `external_source.xlsx` | The workbook the link points **at** (sheet `Data`, `A1 = 42`). |
| `dashboard.xlsx` | Has a **real external link** to `external_source.xlsx`, plus a formula `=[1]Data!A1` that uses it. |

> **Why a generator and not just a committed file?** The external link
> stores an **absolute** `file:///…/external_source.xlsx` path, which is
> only correct on the machine that generated it. If you move this folder,
> re-run `make_example.py` so the link points at the new location. (The
> generator injects the OOXML external-link parts that openpyxl can't
> write on its own — see its docstring for the five pieces involved.)

## How to try it

### Case 1 — link OK (source present)

1. Keep `external_source.xlsx` next to `dashboard.xlsx`.
2. Open **`dashboard.xlsx`** in Excel. (If Excel prompts to update links,
   that's fine — the auditor reads the link list regardless.)
3. Run **Broken External Link Auditor** with **Report** = `All external
   links`.

**Expected:** one external link listed with status **OK** — the source
file exists on disk.

### Case 2 — link BROKEN (source missing)

1. Rename or move `external_source.xlsx` (e.g. to `external_source.bak`).
2. With `dashboard.xlsx` open, run the auditor with **Report** =
   `Broken / offline only`.

**Expected:** the link is reported **BROKEN** — the source path no longer
resolves. Restore the filename to make it OK again.

### Case 3 — OFFLINE (skip the disk probe)

Imagine the source lived on a mapped drive that's currently disconnected.
Put that drive/share prefix in **Offline drives / shares** (e.g. `X:,
\\nas\archive`). Any link whose path starts with one of those is reported
**OFFLINE** instead of BROKEN, and is **not** probed on disk — so the
audit doesn't stall on a long filesystem timeout for an unreachable
server. (Our sample link is a local `D:`/`C:` path, so to see OFFLINE
you'd add this folder's own drive letter to the list.)

## Statuses the auditor reports

| Status | Meaning |
|---|---|
| **OK** | the source file exists |
| **BROKEN** | the source file is missing |
| **URL** | a web/DDE link, not checked on disk |
| **OFFLINE** | path is under a drive/share you flagged offline — reported without probing |

## What this demonstrates

* The tool is **read-only** — it never modifies `dashboard.xlsx`.
* Real external links surfaced via `Workbook.LinkSources`.
* The OK / BROKEN / OFFLINE distinction and the offline-drives timeout
  guard.

> Like every app in this catalog, the auditor is pending live verification
> against a real Excel. If a particular Excel build declines to surface
> the injected link, the scenario above still describes the exact behaviour
> to expect from a workbook with genuine external references.
