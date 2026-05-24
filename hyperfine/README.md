# hyperfine — ScripTree demo

A ScripTree GUI for **[hyperfine](https://github.com/sharkdp/hyperfine)** by [David Peter (sharkdp)](https://github.com/sharkdp).

## What hyperfine does

hyperfine benchmarks one or more shell commands. It runs each command N times, controls for warm-up and per-iteration prepare/cleanup steps, and reports mean / stddev / min / max plus a relative-speed comparison. One-shot CLI — no daemon, no persistent state.

## What's in this demo

| File | Purpose |
|------|---------|
| [`hyperfine.scriptree`](hyperfine.scriptree) | Single-form GUI. ~25 fields across Commands, Run control, Setup/teardown, Parameterization, Naming & comparison, Shell & I/O, Failure handling, Output style, and Export sections. |

The `.scriptree` expects `hyperfine.exe` to live in the same folder. Drop it (or a symlink) next to the file, or edit the `executable` field.

## Repeatable flags

Several hyperfine flags are repeatable (`-p`/`--prepare`, `-C`/`--conclude`, `-L`/`--parameter-list`, `-n`/`--command-name`, `--output`). They're exposed as text fields where you write the full repeated flag for each entry; ScripTree splits the value into multiple argv tokens at run time:

```
--parameter-list compiler gcc,clang --parameter-list opt 0,2,3
```

(See `help/LLM/argument_template.md`'s string-passthrough auto-split rule. v0.1.3+.)

## Why a GUI for a CLI tool?

`hyperfine 'ls' 'fd'` is fine to type. But `hyperfine --warmup 3 --min-runs 50 --prepare 'sync; echo 3 > /proc/sys/vm/drop_caches' --parameter-list compiler gcc,clang --export-json out.json 'gcc -O{opt} foo.c' 'clang -O{opt} foo.c'` benefits hugely from a typed form: numeric inputs for warmup/runs, dropdown for time-unit and style, file pickers for export targets, and named slots so you don't have to remember which flag carries the parameter list vs. the parameter scan.

## Screenshots

### Form view

![Form view of hyperfine](hyperfine_form.png)

### As it appears in the workspace forest

The cell on the right is this demo, docked to the workspace forest hub:

![hyperfine cell docked to the forest](hyperfine_forest.png)

## Installing hyperfine

See the [upstream README](https://github.com/sharkdp/hyperfine#installation). Common paths:

```bash
cargo install hyperfine
# or via package manager:
brew install hyperfine
winget install hyperfine
```

## Upstream

- Repository: <https://github.com/sharkdp/hyperfine>
- Author: [@sharkdp](https://github.com/sharkdp)
- License: see upstream repo

This demo is independent of the hyperfine project; it just generates a GUI front-end for the CLI. Bug reports about hyperfine's behaviour belong upstream; bug reports about the form layout belong here.
