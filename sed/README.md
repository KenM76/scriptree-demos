# sed — ScripTree demo

A ScripTree GUI for **sed**, the classic UNIX stream editor. Implementations:

- **GNU sed** — Default on Linux distros, available everywhere via package managers. Most permissive. **The form's flag descriptions assume GNU sed** (`-i` without a mandatory suffix arg, `-s`, `-z`, `--regexp-extended`).
- **BSD sed** — Default on macOS and the BSDs. Slightly different `-i` semantics (requires an explicit backup-suffix argument, even if empty).

## What sed does

Stream editor. Reads input one line at a time, runs a small script of edit commands against each line, prints the result. The classic use is find-and-replace:

```bash
sed 's/foo/bar/g' input.txt
```

But sed also handles deletion, insertion, conditional printing, multi-line transforms, and (with care) some pretty fancy text wrangling. It's the spiritual ancestor of `s///` in Perl and Vim.

## When to reach for sed

- **Substitution across files** — `s/old/new/g`, optionally with `-i` to overwrite the originals.
- **Stripping noise** — leading whitespace, trailing carriage returns, blank lines.
- **Extracting matches** — `-n '/pattern/p'` works as a grep that also lets you transform.
- **Line-number selections** — print/delete lines 1-10, the last line, lines between two markers.

When NOT to reach for sed:

- Anything multi-line that isn't a simple "delete blank lines" — sed CAN do it but the syntax is painful. awk is friendlier.
- Anything stateful or computed. Use awk or Python.
- Find-and-replace inside structured formats (JSON, XML, YAML). Use a parser.

## Glossary

- **Script** — The whole sed program: one or more commands separated by semicolons or newlines.
- **Command** — A single edit operation: `s` (substitute), `d` (delete), `p` (print), `i\` (insert), `a\` (append), `c\` (change), `q` (quit), and a few more.
- **Address** — Optional prefix on a command saying which lines it applies to: `5` (line 5), `$` (last line), `1,5` (lines 1-5), `/regex/` (lines matching regex).
- **Pattern space** — sed's internal buffer holding the current line being processed. The script's commands edit this buffer; auto-print writes it out.
- **Hold space** — A second internal buffer you can stash text in (with `h`/`H`) and retrieve later (with `g`/`G`). Used for multi-line transforms.
- **`s/regex/replacement/flags`** — The substitute command. Common flags: `g` (global — replace all matches on the line), `i` (case-insensitive, GNU only), `p` (print after substitution), a number (replace just the Nth match).
- **Basic vs. Extended regex** — sed defaults to "basic" regex where `+ ? ( ) |` are LITERAL and you backslash-escape them to make them special. Tick `-E` to flip that — `+ ? ( ) |` become special and you escape them to make them literal. Modern users almost always want `-E`.

## Worked examples

```sed
# Basic find/replace (first match per line):
s/foo/bar/

# Find/replace ALL matches per line:
s/foo/bar/g

# Case-insensitive (GNU):
s/foo/bar/gi

# Use a different delimiter when '/' is in the pattern:
s|/path/to/old|/path/to/new|g

# Delete all lines matching a pattern:
/^DEBUG/d

# Delete blank lines:
/^$/d

# Delete the first 10 lines:
1,10d

# Print just lines 5-10:
-n '5,10p'        # (with -n flag)

# grep clone — print only matching lines:
-n '/error/p'     # (with -n flag)

# Strip leading whitespace:
s/^[[:space:]]*//

# Strip Windows line endings:
s/\r$//

# Two commands at once (semicolon-separated):
s/foo/bar/g; s/baz/qux/g
```

## What's in this demo

| File | Purpose |
|------|---------|
| [`sed.scriptree`](sed.scriptree) | Single-form GUI. ~9 fields covering the script (or script file), input files, and the common mode flags (`-n`, `-E`, `-i`, `-s`, `-z`). |

The form expects `sed` to be on `PATH`. Almost universally available — Git Bash on Windows bundles GNU sed; macOS ships BSD sed (install GNU via `brew install gnu-sed` if you want consistent flag behaviour with the form's descriptions).

## Why a GUI for a one-liner tool?

Honestly, simple sed one-liners are faster typed at the terminal. The form earns its keep when:

- The script is more than one line and shell-quoting gets painful.
- You're using `-i` (in-place edit) and want to be REALLY sure you haven't typo'd the regex before you overwrite the files.
- You want to save a "this transformation against this set of files" combo as a named configuration to re-run later.

## Installing sed

```bash
# Already installed? Try:
sed --version

# macOS (GNU sed, recommended for cross-platform scripts):
brew install gnu-sed
# adds 'gsed' — change the form's executable to 'gsed' if you want GNU semantics

# Debian / Ubuntu:
sudo apt install sed

# Fedora / RHEL:
sudo dnf install sed

# Windows: Git for Windows bundles GNU sed, or:
choco install sed
# or:
scoop install sed
```

## Upstream

- GNU sed: <https://www.gnu.org/software/sed/>
- BSD sed: ships with macOS / FreeBSD / OpenBSD / NetBSD base systems

This demo wraps the standard `sed` invocation; the form-level flags are the cross-implementation ones with GNU-only flags marked as such in their descriptions.
