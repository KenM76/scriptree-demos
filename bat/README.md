# bat — ScripTree demo

A ScripTree GUI for **[bat](https://github.com/sharkdp/bat)** by [David Peter (sharkdp)](https://github.com/sharkdp).

## What bat does

bat is a `cat` clone with syntax highlighting, Git integration, line numbers, and an automatic pager. One-shot CLI: every flag is a presentation knob (theme, language, style components, line ranges, wrapping, paging).

## What's in this demo

| File | Purpose |
|------|---------|
| [`bat.scriptree`](bat.scriptree) | Single-form GUI. ~33 fields across Input, Highlighting, Display, Layout, Color & paging, and Diff & misc sections. |

The `.scriptree` expects `bat.exe` to live in the same folder. Drop it (or a symlink) next to the file, or edit the `executable` field.

## Repeatable flag caveat

`--map-syntax`, `-r`/`--line-range`, and `-H`/`--highlight-line` accept multiple values upstream. ScripTree's current schema has no first-class repeatable-flag widget, so the first instance is a typed field; for additional `--map-syntax` entries there's a paired raw-text field where you write the full flag for each. Example:

```
--map-syntax '*.conf:INI' --map-syntax '.envrc:Bash'
```

For multiple line-range or highlight-line entries, comma-separate inside the field or call bat directly.

## Why a GUI for a CLI tool?

Plain `bat foo.rs` is faster typed than form-filled — that's the whole point. The GUI earns its keep when you start composing `--style numbers,changes --theme TwoDark --line-range 100:150 --highlight-line 110:115 --map-syntax '*.tf:HCL' --paging never`. Six knobs, six widgets, no man-page round-trip.

## Installing bat

See the [upstream README](https://github.com/sharkdp/bat#installation). Common paths:

```bash
cargo install --locked bat
# or via package manager:
brew install bat
winget install sharkdp.bat
```

## Upstream

- Repository: <https://github.com/sharkdp/bat>
- Author: [@sharkdp](https://github.com/sharkdp)
- License: see upstream repo

This demo is independent of the bat project; it just generates a GUI front-end for the CLI. Bug reports about bat's behaviour belong upstream; bug reports about the form layout belong here.
