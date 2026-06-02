# Broken External Link Auditor (Excel)

Audit the **external workbook links** of the Excel workbook that is open and
active, and report each link's source as **OK** (file exists), **BROKEN** (file
missing), **URL** (a web/DDE link not checked on disk), or **OFFLINE** (lives on
a drive/share you flagged as unavailable, so it is reported without a slow
filesystem timeout). Use it to find dead `='C:\...\[Book.xlsx]Sheet'!A1`-style
references before they silently feed `#REF!` into a model.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

---

## 1. What the user sees (end-user guide)

### The form fields

| Field | Meaning |
|---|---|
| **Report** (default `Broken / offline only`) | `Broken / offline only` lists just the problem links (what you usually want). `All external links` lists every external link with its status — a full inventory. |
| **Offline drives / shares (optional)** | Comma-separated drive letters or UNC prefixes that are currently disconnected (e.g. `X:, \\nas\archive`). Any link source whose path **starts with** one of these is reported `OFFLINE` and **not probed on disk**, avoiding a long filesystem timeout when a mapped drive or server is unreachable. Blank = probe every link normally. |
| **Output format** | `Markdown table` (paste into a ticket) or `Plain text` (aligned list). |

### What it does to your workbook

**Nothing.** It is strictly read-only — it never edits, breaks, redirects, or
re-links anything. It only reads the workbook's link-source list and checks the
filesystem.

### Prerequisites

* Excel is **running with the target workbook open and active**.
* ScripTree's bundled `lib/combridge/combridge.exe` is present.

---

## 2. The logic (reconstruct-the-tool spec)

A **Strategy-A shim**: a ScripTree form runs `link_auditor.py`; the shim bakes
the form values into a C# Roslyn script rendered from
`link_auditor.csx.template`, then runs it through combridge, which owns the live
COM connection to Excel.

* **combridge `run-script` has no argv channel** — a `.csx` only sees the plugin
  globals (`xlApp` / `xlBook` / `xlSheet`) and environment. Form values are baked
  in by replacing `__TOKEN__` placeholders.
* **combridge swallows the script's `return` value** — clean run exits 0. So the
  script prints a first-line **sentinel** and the shim translates it into the
  process exit code.

### Sentinel + status → exit-code contract

First stdout line of the `.csx`:

```
__LINKAUDIT__ STATUS=<code> [key=value ...]
```

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Audit ran; report follows (`broken=B offline=F`). |
| `NO_WORKBOOK` | 2 | No workbook open/active in Excel. |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes. |

### What the `.csx` does

1. **Guard:** `xlBook is null` → `NO_WORKBOOK`.
2. **Parse offline prefixes:** split the baked `offlineInput` on `,`, trim, drop
   blanks. `StartsWithAny` compares case-insensitively (`OrdinalIgnoreCase`).
3. **Enumerate links:** `xlBook.LinkSources(XlLink.xlExcelLinks)` (`xlExcelLinks
   = 1`) returns a 1-based `System.Array` of **full file-path strings**, or
   **null** when the workbook has no such links. Guard with `raw is System.Array
   arr`; skip empty entries.
4. **Classify each source** (order matters):
   * contains `"://"` → `URL` (web/DDE link — never touch the disk).
   * else starts with an offline prefix → `OFFLINE` (reported without probing).
   * else `File.Exists(src)` → `OK`.
   * else → `BROKEN`.
5. **Tally** OK / BROKEN / OFFLINE / URL counts. `problemsOnly = reportMode !=
   "all"`; a problem is `BROKEN` or `OFFLINE`. Build the `shown` list (all rows
   in `all` mode; only problems otherwise).
6. **Report:** sentinel `__LINKAUDIT__ STATUS=OK broken={B} offline={F}` then:
   * `total == 0` → "no external workbook links found".
   * `problemsOnly && shown.Count == 0` → "no broken or offline links found"
     plus the summary line.
   * otherwise a markdown table (`| Status | Link source |`, pipes in paths
     escaped `\|`) or an aligned plain-text list (`[STATUS ] <source>`), ending
     with a one-line summary: `{total} external links - {ok} OK, {broken}
     broken, {offline} offline, {url} URL.`

### Excel COM facts this relies on (see the office-com RAG)

* **`Workbook.LinkSources(xlExcelLinks=1)`** returns a 1-based `System.Array` of
  **full path strings** (not the `[bracketed]` display form), or **null** when
  none. `File.Exists(src)` works directly on each.
* `using System.IO;` (for `File.Exists`) **is** in combridge's default Excel
  usings — no need to add it (unlike the Outlook plugin).
* This is a read-only path, so it is **safe from the modal-dialog hang** that
  afflicts COM-launched Excel write/save/open paths — no `DisplayAlerts` /
  `AskToUpdateLinks` juggling needed here.

---

## 3. Files in this app

| File | Role |
|---|---|
| `link-auditor.scriptree` | The form (3 params; `search` icon embedded as PNG — `search` is reserved for read-only auditors). |
| `link-auditor.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config. |
| `link_auditor.py` | The Strategy-A shim. |
| `link_auditor.csx.template` | The Roslyn template with `__TOKEN__` placeholders. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
link_auditor.py
  --report-mode broken|all
  --offline-drives <str>        (comma-separated prefixes; may be empty)
  --output-format markdown|text
```

All three pass through `["--flag","{id}"]` token groups; an empty
`offline_drives` drops the value but keeps `--offline-drives` (argparse default
`""` handles it).

---

## 4. Editing / maintenance notes

* **Keep it read-only.** The whole value proposition is a safe audit — never add
  a "fix/relink" write path here without re-opening the modal-hang question (and
  the `DisplayAlerts` guards it would require).
* **Classification order is load-bearing:** URL before OFFLINE before the
  `File.Exists` probe. The OFFLINE short-circuit exists specifically to skip the
  probe, so it must precede it.
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree.cli.validate <path>`.
* **Offline render-check:** render the template with sample values (paths with
  backslashes, an offline prefix, a URL) and grep for `__[A-Z_]+__` — only the
  `__LINKAUDIT__` sentinel should match; anything else is an unfilled placeholder.
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path.
