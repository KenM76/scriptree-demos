# Mailbox Age & Size Auditor (Outlook)

Scan your Outlook stores and report, **per folder**, the item count, total size,
the age span (oldest → newest item), and a "dead weight" count + size of items
older than a chosen number of years. Use it to find what is bloating your mailbox
before you hit a quota.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

---

## 1. What the user sees (end-user guide)

### The form fields

| Field | Meaning |
|---|---|
| **Scope** | `All stores` (default) audits your live mailbox **plus every attached PST archive** — the edge over Outlook's built-in Mailbox Cleanup, which ignores PSTs. `Default mailbox only` restricts to your primary account. |
| **Flag items older than (years)** (default 2) | Age threshold for the "dead weight" counter. Per folder, the report shows how many items — and how much space — are older than this. `0` counts every item as stale. Range 0–100. |
| **Hide folders smaller than (MB)** (default 5) | Folders below this size are omitted, hiding the dozens of empty system folders (Sync Issues, Quick Step Settings, …) that clutter every profile. `0` shows every folder. Range 0–1,000,000. |
| **Output format** | `Markdown table` (paste into a ticket) or `Plain text`. |

### What it does to your mailbox

**Nothing.** It is strictly read-only — it never moves, deletes, or modifies a
single item. It only reads aggregate stats.

### Prerequisites

* Outlook is **running with a mailbox loaded**.
* ScripTree's bundled `lib/combridge/combridge.exe` is present.

---

## 2. The logic (reconstruct-the-tool spec)

A **Strategy-A shim**: a ScripTree form runs `mailbox_auditor.py`; the shim bakes
the form values into a C# Roslyn script rendered from `mailbox_auditor.csx.template`,
then runs it through combridge, which owns the live COM connection to Outlook.

* **combridge `run-script` has no argv channel** — a `.csx` only sees the plugin
  globals (`olApp` / `olNs` / `olExplorer`) and environment. Form values are baked
  in by replacing `__TOKEN__` placeholders.
* **combridge swallows the script's `return` value** — clean run exits 0. So the
  script prints a first-line **sentinel** and the shim translates it into the
  process exit code.

### Sentinel + status → exit-code contract

```
__MAILAGE__ STATUS=<code> [key=value ...]
```

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Audit ran; report follows (`folders=N stores=S`). |
| `NO_STORE` | 2 | No Outlook stores found (no session loaded). |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes. |

### What the `.csx` does

1. **Choose stores.** `default` scope → `olNs.DefaultStore` (the primary account);
   `all` scope → every `olNs.Stores` member (live mailbox **and** attached PSTs).
   No stores → `NO_STORE`.
2. **Walk each store's folder tree** recursively from `store.GetRootFolder()`,
   building one record per folder.
3. **Per-folder stats via `Folder.GetTable` — never iterating `MailItem`.** On a
   folder with tens of thousands of items, looping `MailItem` freezes Outlook. The
   fast path: `folder.GetTable("", olUserItems)`, `Columns.RemoveAll()`, then
   `Columns.Add(PR_MESSAGE_SIZE)` (`0x0E080003`) and `Columns.Add("ReceivedTime")`,
   and sum the rows. `Row[...]` is **1-based**; `ReceivedTime` is **null on
   non-mail items** (guard with `is DateTime`). The age cutoff is
   `DateTime.Now.AddYears(-staleYears)`; an item older than the cutoff adds to the
   folder's `staleCount`/`staleBytes`. Some search/system folders **throw** on
   `GetTable` — wrap it (and the `folder.Folders` enumeration) in try/catch so one
   stubborn folder doesn't abort the walk.
4. **Filter + sort for display:** keep folders at/above `minFolderMb`, sort
   largest-first; count the hidden remainder for the summary line.
5. **Report:** sentinel `__MAILAGE__ STATUS=OK folders={shown} stores={S}` then the
   markdown/plain-text report (per-folder size / item count / age span / stale
   count+size, plus grand totals and how many small folders were hidden).

### Outlook COM facts this relies on (see the office-com RAG)

* **`Folder.GetTable` is the only safe way to aggregate** — never loop `MailItem`
  for stats. `PR_MESSAGE_SIZE` (`0x0E080003`, PT_LONG → boxed int) includes
  attachments; there is no folder-level total-size property to read directly.
* `olNs.Stores` enumerates every account + attached PST; `DefaultStore` is just
  the primary. `store.GetRootFolder()` (cast to `Folder`) is the tree root.
* Read-only audits are safe from the modal-dialog hang that afflicts write paths.

---

## 3. Files in this app

| File | Role |
|---|---|
| `mailbox-auditor.scriptree` | The form (4 params; `search` icon embedded as PNG — `search` is reserved for read-only auditors). |
| `mailbox-auditor.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config. |
| `mailbox_auditor.py` | The Strategy-A shim. |
| `mailbox_auditor.csx.template` | The Roslyn template with `__TOKEN__` placeholders. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
mailbox_auditor.py
  --scope all|default
  --stale-years <int>           (clamped >= 0 in the shim)
  --min-folder-mb <int>         (clamped >= 0 in the shim)
  --output-format markdown|text
```

All four pass through `["--flag","{id}"]` token groups.

---

## 4. Editing / maintenance notes

* **Never replace `GetTable` with a `MailItem` loop** — that's the whole point of
  the design (see the RAG lesson). Keep the per-folder and per-enumeration
  try/catch guards.
* **The integer `min`/`max` bounds** (`stale_years` 0–100, `min_folder_mb`
  0–1,000,000) are hand-edited fields that an icon re-embed would strip — re-add
  them as the last edit after any `embed_icon` round-trip. The shim also clamps
  both to `>= 0`.
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree.cli.validate <path>`.
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path.
