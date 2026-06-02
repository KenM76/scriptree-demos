# Example — Mailbox Age & Size Auditor

## Why there's no sample file here

Unlike the Word / Excel / PowerPoint tools, this app takes **no input file**.
It reads your **live Outlook profile** — every store (mailbox) and any attached
`.pst` archive — straight through MAPI. There is nothing to hand you as a
`sample.xxx`; the "input" is whatever Outlook has open on your machine. So this
folder is a **scenario walkthrough** instead: a described mailbox, the exact
form values to use, and the report you'd get back.

The tool is **strictly read-only** — it never moves, deletes, or modifies a
single item. It only *reads* per-folder statistics.

## The scenario

Imagine an Outlook profile with two stores:

* **`Ken@example.com`** — the live Exchange/IMAP mailbox.
* **`Archive 2019-2021.pst`** — an old PST the user attached years ago and
  forgot about. (The built-in *Mailbox Cleanup* tool ignores PSTs entirely —
  surfacing this is half the point of the app.)

Rough contents:

| Folder | Items | Size | Oldest | Newest |
|---|--:|--:|---|---|
| `Ken@example.com\Inbox` | 9,120 | 1.4 GB | 2018-03 | 2026-05 |
| `Ken@example.com\Sent Items` | 6,400 | 980 MB | 2018-03 | 2026-05 |
| `Ken@example.com\Inbox\Newsletters` | 12,300 | 320 MB | 2019-01 | 2026-04 |
| `Ken@example.com\Deleted Items` | 540 | 88 MB | 2024-02 | 2026-05 |
| `Ken@example.com\Drafts` | 11 | 2 MB | 2025-11 | 2026-05 |
| `Archive 2019-2021.pst\Inbox` | 22,000 | 3.1 GB | 2019-01 | 2021-12 |

## How to try it (on your own mailbox)

1. Make sure **Outlook is running** with your mailbox loaded.
2. Run **Mailbox Age & Size Auditor** from ScripTree.

| Field | Value | Effect |
|---|---|---|
| Scope | `all` | audits every store **including PSTs** (use `default` for just the primary mailbox) |
| Flag items older than (years) | `2` | items older than this count as "dead weight" |
| Hide folders smaller than (MB) | `5` | tiny system folders are dropped from the table to keep it high-signal |
| Output format | `markdown` | a rendered table (use `text` for a plain monospace report) |

## Expected output (markdown mode)

Folders are sorted **largest-first**; anything under the 5 MB floor (here,
`Drafts`) is hidden but still counted in the grand totals. With the 2-year
threshold, "Older than 2 years" counts items received before ~2024-05.

```
# Mailbox age & size audit

| Folder | Items | Size | Age span | Older than 2 years |
|---|--:|--:|---|---|
| Archive 2019-2021.pst\Inbox | 22,000 | 3.10 GB | 2019-01 → 2021-12 | 22,000 (3.10 GB) |
| Ken@example.com\Inbox | 9,120 | 1.4 GB | 2018-03 → 2026-05 | 4,300 (612 MB) |
| Ken@example.com\Sent Items | 6,400 | 980 MB | 2018-03 → 2026-05 | 3,900 (520 MB) |
| Ken@example.com\Inbox\Newsletters | 12,300 | 320 MB | 2019-01 → 2026-04 | 7,100 (180 MB) |
| Ken@example.com\Deleted Items | 540 | 88 MB | 2024-02 → 2026-05 | 60 (9 MB) |

**Summary:** Audited 2 stores, 6 folders (50,371 items, 5.85 GB total). Older than 2 years: 37,360 items, 4.30 GB. 1 folder below 5 MB hidden.
```

(The numbers above are illustrative; your run reports your real folders.)

## How to read it

* **The entire PST is dead weight** — every item predates the 2-year cutoff.
  That's 3.1 GB you could detach/archive offline. The native cleanup tool
  would never have shown you this.
* **Inbox + Sent** carry large old tails (the `Older than 2 years` column),
  even though they're still active (newest item this month) — candidates for
  an archive-and-purge of the old portion.
* **Newsletters** is huge by *count* (12k) but modest by size — a different
  kind of clutter.

## What this demonstrates

* Per-folder item count, total size, and oldest→newest **age span**.
* The **dead-weight** column: count + size of items older than your threshold.
* Auditing **PSTs as well as the live mailbox** (the differentiator vs. the
  built-in tool).
* Largest-first sorting and the size-floor filter for a high-signal report.
* Strictly **read-only** — nothing in the mailbox is touched.

> The tool reads lightweight MAPI tables (message size + received time) rather
> than opening each message, so it stays fast even on 40k-item folders. Pending
> live verification against a real Outlook profile; the scenario above
> describes the exact report shape to expect.
