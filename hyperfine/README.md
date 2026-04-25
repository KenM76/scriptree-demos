# hyperfine — ScripTree demo

A ScripTree GUI for **[hyperfine](https://github.com/sharkdp/hyperfine)** by [David Peter (sharkdp)](https://github.com/sharkdp).

## What hyperfine does

hyperfine benchmarks one or more shell commands. It runs each command N times, controls for warm-up and per-iteration prepare/cleanup steps, and reports mean / stddev / min / max plus a relative-speed comparison. One-shot CLI — no daemon, no persistent state.

## What's in this demo

| File | Purpose |
|------|---------|
| [`hyperfine.scriptree`](hyperfine.scriptree) | Single-form GUI. ~25 fields across Commands, Run control, Setup/teardown, Parameterization, Naming & comparison, Shell & I/O, Failure handling, Output style, and Export sections. |

The `.scriptree` expects `hyperfine.exe` to live in the same folder. Drop it (or a symlink) next to the file, or edit the `executable` field.

## Repeatable flag caveat

Several hyperfine flags are repeatable (`-p`/`--prepare`, `-C`/`--conclude`, `-L`/`--parameter-list`, `-n`/`--command-name`, `--output`). ScripTree's current schema has no first-class repeatable-flag widget, so these are exposed as raw text fields where you write the full flag for each entry, e.g.:

```
--parameter-list compiler gcc,clang --parameter-list opt 0,2,3
```

If/when ScripTree grows a list widget, those fields are obvious upgrade targets.

## Why a GUI for a CLI tool?

`hyperfine 'ls' 'fd'` is fine to type. But `hyperfine --warmup 3 --min-runs 50 --prepare 'sync; echo 3 > /proc/sys/vm/drop_caches' --parameter-list compiler gcc,clang --export-json out.json 'gcc -O{opt} foo.c' 'clang -O{opt} foo.c'` benefits hugely from a typed form: numeric inputs for warmup/runs, dropdown for time-unit and style, file pickers for export targets, and named slots so you don't have to remember which flag carries the parameter list vs. the parameter scan.

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
