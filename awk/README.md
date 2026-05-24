# awk — ScripTree demo

A ScripTree GUI for **awk**, the classic UNIX text-processing language by Aho, Weinberger, and Kernighan (1977). Most modern systems ship one of three implementations:

- **gawk** — GNU awk. The most feature-rich; default on Linux distros, available everywhere via package managers. **The form's flag descriptions assume gawk** (the `--posix`, `--traditional`, `-b` flags are gawk-only).
- **mawk** — A faster but more minimal implementation. Sometimes the default on Debian/Ubuntu.
- **BSD awk** (a.k.a. nawk, "one true awk") — The default on macOS and the BSDs.

## What awk does

You feed awk lines of text. For every line, awk runs a tiny program you write. The program is a list of `pattern { action }` pairs — for each input line, every pattern that matches triggers its action. That's it. The whole language is built on top of that idea.

It's the right tool when you have:

- Tabular data (CSV, TSV, log files, /etc/passwd-style colon-separated files) and want to extract / sum / filter columns.
- A pile of similar lines and you want to group / count / reformat them.
- A dataset that's too small for a database but too tabular for grep alone.

It's the *wrong* tool when:

- You need real data structures (use Python).
- The work is bigger than one terminal session (write a script in any real language).
- You need anything multi-line or stateful that's awkward in awk's pattern-action model — sed and awk both prefer line-at-a-time.

## What you might know already (or not)

If awk is new, here's the conceptual ladder:

1. **A program is just `{ print }`** — by default awk prints every line that matched. No pattern means "every line", no action means `{ print }`.
2. **Awk auto-splits each line** into `$1, $2, ...` — fields separated by whitespace by default. So `{ print $1 }` prints the first whitespace-separated word of every line.
3. **A pattern can be a regex** like `/error/` (matches lines containing 'error') or an expression like `NF > 0` (matches lines with at least one field). Lines where the pattern is true → run the action.
4. **`BEGIN { ... }`** runs once before any input. **`END { ... }`** runs once after all input. That's how you do "set up a counter, count lines, print at the end".
5. **Variables don't need declaring** and start as `0` or `""`. So `{ sum += $1 } END { print sum }` just works.

That's 90% of what you'll ever need. The man page has the rest.

## Glossary

- **Field** — A piece of a line, separated by the field separator (whitespace by default). Available as `$1`, `$2`, ..., with `$NF` being the last and `$0` being the whole line.
- **Record** — Awk's word for "input line" (because in theory the record separator could be something other than newline; `RS` controls it).
- **Pattern** — The condition before the `{ action }`. A regex, an expression, the literal `BEGIN`, `END`, or empty (matches every line).
- **Action** — Code in `{ ... }`. Runs when the pattern matches.
- **`NR`** — Number of records read so far (current line number, 1-based across all files).
- **`NF`** — Number of fields on the current line.
- **`FS`** — Input field separator. Set with `-F` from the command line, or `BEGIN { FS = "," }` in the program.
- **`OFS`** — Output field separator (used when you `print $1, $2, $3` — the commas become OFS).
- **`FILENAME`** — Name of the current input file.
- **`gsub(regex, replacement, target)`** — Global substitute. Replaces every match of regex in target with replacement. `target` defaults to `$0` if omitted.
- **`sub(regex, replacement, target)`** — Same as gsub but only replaces the first match.
- **`split(string, array, separator)`** — Splits a string into an array.
- **`printf "fmt", args...`** — C-style formatted output. Use this when you want columns aligned.
- **`print` vs. `printf`** — `print` adds a newline; `printf` doesn't (you add `\n` yourself).

## Worked examples

The form's `Program` field documents these too. Reproduced here for skim-readability:

```awk
# Print the first column of a CSV
{ print $1 }

# Sum the third column
{ sum += $3 } END { print sum }

# Average the first column (use NR for the count)
{ s += $1 } END { print s / NR }

# Print only lines containing 'error'
/error/ { print }

# Print only non-blank lines
NF > 0

# Number every line
{ print NR, $0 }

# Print just lines 10 through 20
NR >= 10 && NR <= 20

# Print the LAST field of each line
{ print $NF }

# Count distinct values of column 1
{ count[$1]++ } END { for (k in count) print k, count[k] }

# Convert space-separated to TSV
BEGIN { OFS = "\t" } { $1 = $1; print }
```

## What's in this demo

| File | Purpose |
|------|---------|
| [`awk.scriptree`](awk.scriptree) | Single-form GUI. ~8 fields covering the program (or program file), input files, field separator, a `-v` variable assignment, and gawk's compatibility flags. |

The form expects `awk` to be on `PATH`. On Linux that's a given. On macOS, the system awk is BSD awk; install gawk via `brew install gawk` if you want the GNU extensions (the form's compat flags assume gawk). On Windows, the easiest path is Git Bash (it bundles a gawk), or install gawk standalone via Chocolatey / Scoop / WSL.

## Why a GUI for a language?

Honestly, awk one-liners are usually faster typed at the terminal. The form earns its keep when:

- The program is more than one line (the textarea is more pleasant than escape-quoted shell strings).
- You're using `-F` with a tricky separator and want a labelled field rather than counting backslashes.
- You're flipping between programs against the same data and want to save them as named configurations.

## Screenshots

### Form view

![Form view of awk](awk_form.png)

### As it appears in the workspace forest

The cell on the right is this demo, docked to the workspace forest hub:

![awk cell docked to the forest](awk_forest.png)

## Installing awk

```bash
# Already installed? Try:
awk --version

# macOS (gawk):
brew install gawk

# Debian / Ubuntu (gawk or mawk, both fine):
sudo apt install gawk

# Fedora / RHEL:
sudo dnf install gawk

# Windows: install Git for Windows (bundles gawk), or:
choco install gawk
# or:
scoop install gawk
```

## Upstream

- gawk (GNU awk): <https://www.gnu.org/software/gawk/>
- BSD / "one true awk": <https://github.com/onetrueawk/awk>
- mawk: <https://invisible-island.net/mawk/>

This demo wraps the standard `awk` invocation; it's not specific to any one implementation. The form-level flags are the cross-implementation ones; the gawk-only flags (`--posix`, `--traditional`, `-b`) are flagged as such in the descriptions.
