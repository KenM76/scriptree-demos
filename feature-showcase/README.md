# feature-showcase

A teaching demo that exercises every widget and every important
feature of the ScripTree `.scriptree` schema in one form.

If you are new to ScripTree, **start here**. Click through every tab,
read each field's description, then open
[`feature-showcase.scriptree`](./feature-showcase.scriptree) in a text
editor and compare the rendered widget with the JSON behind it.

## What's in the form

| Tab | What it teaches |
|-----|-----------------|
| **Browse** | `radio` (Mode picker), `folder` with drag-and-drop, `text` (Name filter), preset-bundle `dropdown` (Preset filter), dynamic `dropdown` populated by `choices_provider` (Preview file), `checkbox` (Recurse / Include hidden), `number` with `min`/`max`/`step` (Max depth), `visible_when` gating fields by mode. |
| **Output** | `checkbox_list` `multiselect` with `select_all` (Columns), `folder_list` with `must_exist` (Search folders), `file_list` with `file_filter` (Reference files). |
| **Safety** | `checkbox` showing the destructive-intent convention real tools use, `textarea` with drag-and-drop. |
| **About this demo** | `text` (URL / masking-heuristic demo), `save_file` with `file_filter`. |

The cell appearance (icon, fill colour, text label) is set in the
`cell` sub-object at the bottom of the JSON — embedded PNG, custom
hex fill, white text label.

## What's in the argv

The argument_template builds a Windows PowerShell pipeline:

```
powershell -NoLogo -NoProfile -NonInteractive -Command \
  Get-ChildItem -Path <target_dir> \
    [-Force] [-Recurse] [-Depth N] [-Filter <pattern>] \
    <preset_filter> \
    [-Include <file> ...]  (one per entry, fanned out) \
    [-LiteralPath <folder> ...]  (one per entry, fanned out) \
  | Select-Object <col1> <col2> ...  (comma-joined) \
  | Format-Table -AutoSize
```

Two patterns to study in the template:

1. **Token-group fan-out for `folder_list` / `file_list`:** the entries
   `["-LiteralPath", "{search_folders}"]` and `["-Include",
   "{reference_files}"]` repeat themselves once per selected element.
   With three folders chosen, the argv contains
   `-LiteralPath A -LiteralPath B -LiteralPath C`. See
   `argument_template.md` §5.

2. **Comma-join for the bare `{columns}` placeholder:** the
   checkbox_list's selected columns are emitted as a single comma-joined
   token, which Select-Object happens to accept (PowerShell parses the
   comma-list internally). Compare with the fan-out pattern above to
   see how a single line in the template controls argv shape.

## Cross-references to canonical docs

Everything this demo does is documented in detail elsewhere:

- Widget × type matrix, `radio` vs `checkbox`, masking heuristic, the
  preset-bundle pattern, `folder_list` / `file_list` field shape →
  `D:/Dev/ScripTree/docs/LLM/param_types_widgets.md`
- Argument template grammar, token groups, fan-out vs comma-join,
  conditional flags, repeatable-flag pattern →
  `D:/Dev/ScripTree/docs/LLM/argument_template.md`
- `choices_provider`, `depends_on`, the JSON contract a provider
  script has to honour, refresh modes →
  `D:/Dev/ScripTree/docs/LLM/dynamic_providers.md`
- `visible_when` / `required_when` grammar, section layouts, cell
  cosmetics, `confirmation` field →
  `D:/Dev/ScripTree/docs/LLM/scriptree_format.md`
- Icon set, picking rules, embed workflow →
  `D:/Dev/ScripTree/docs/LLM/icon_library.md`

## Files

| File | Purpose |
|------|---------|
| `feature-showcase.scriptree` | The form. |
| `list_files.py` | Tiny `choices_provider` script. Reads JSON from stdin, prints a `{choices, choice_labels, default}` JSON document to stdout. Run by ScripTree at form-open and whenever 'Target directory' changes. |
| `README.md` | This file. |

## How to run

1. Open `feature-showcase.scriptree` in ScripTree (V1 editor / runner,
   V3 cell shell, or as a standalone window).
2. Pick a Target directory (default is the demo folder itself).
3. Click around. Watch how `Mode` changes the visible fields. Tick
   columns, drag folders onto Search folders, then hit Run.
4. The PowerShell command preview shows you the assembled argv before
   it spawns — read it to see how the template grammar produced it.

No upstream tool to install: PowerShell is part of Windows.

## Screenshots

### Form view

![Form view of feature-showcase](feature-showcase_form.png)

### As it appears in the workspace forest

The cell on the right is this demo, docked to the workspace forest hub:

![feature-showcase cell docked to the forest](feature-showcase_forest.png)
