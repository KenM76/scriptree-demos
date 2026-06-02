#!/usr/bin/env python3
"""Column-Based Sheet Segregator (Excel) — ScripTree shim (Strategy A).

================================================================================
WHAT THIS PROGRAM DOES (the logic — read this to understand/rewrite the tool)
================================================================================
The user has a flat table on the active Excel worksheet (rows of data, usually
with a header row). They want to split that table into separate sheets — one
sheet per distinct value found in a chosen "key" column. Classic example: a
sales export with a "Region" column → produce one sheet per region, each holding
only that region's rows (plus a copy of the header row).

The whole operation is **non-destructive to the source data**:

  * In ``new_workbook`` mode (the default, safest) the tool creates a brand-new
    workbook and writes the per-value sheets there. The original workbook is
    never modified at all.
  * In ``new_sheets`` mode the tool adds the per-value sheets to the *current*
    workbook (after the active sheet). The source sheet itself is read-only;
    only new sheets are added. A "backup the workbook first" guard is offered
    because adding sheets changes the workbook structure.

In both modes the result is left **open and UNSAVED** in Excel so the user can
review it and save (or discard) it themselves. The tool never silently writes a
file over anything the user already has on disk — the single exception is the
optional backup copy in ``new_sheets`` mode, which is written to a *new* path
(``<name>_backup.<ext>``) and never overwrites the working file.

================================================================================
WHY THIS IS A SHIM (Strategy A — the project-wide integration pattern)
================================================================================
ScripTree is a GUI form that runs a command line. It launches this script with
``executable = python`` and passes the form field values as argv (see the
``argument_template`` in ``sheet-segregator.scriptree``). The actual Excel work
must happen inside combridge (which owns the COM connection to the running
Excel), and combridge's ``run-script`` subcommand executes a C# Roslyn script
(``.csx``).

combridge's ``run-script`` has **no argv channel** — a ``.csx`` can only see the
plugin globals (``xlApp``/``xlBook``/``xlSheet``) plus environment variables. So
this shim BAKES the form values into a generated ``.csx`` (rendered from
``sheet_segregator.csx.template`` by substituting ``__PLACEHOLDER__`` tokens),
then hands that script to::

    combridge.exe excel run-script <temp.csx> -

================================================================================
WHY THE SENTINEL + SHIM-OWNED EXIT CODE (a hard combridge constraint)
================================================================================
combridge's ScriptHost **ignores the C# script's ``return`` value** — it exits 0
on any clean run (only a compile error = 3, an unhandled throw = 4, or a host
error = 5 produce a non-zero code). ScripTree, however, decides success/failure
from the child process exit code. So we use a two-part protocol:

  1. The ``.csx`` prints a machine-readable **sentinel** as its first stdout
     line:  ``__XLSEG__ STATUS=<code> [key=value ...]``  followed by the
     human-readable report.
  2. THIS shim parses that sentinel, strips it from the output, prints the
     report, and **translates the status into the process exit code** that
     ScripTree sees.

Status → exit-code contract (see ``main`` and the ``.csx`` template):

    STATUS=OK        -> exit 0   (segregation succeeded; report follows)
    STATUS=NOWB      -> exit 2   (no workbook open)
    STATUS=NODATA    -> exit 2   (active sheet has no usable table)
    STATUS=BADCOL    -> exit 2   (key column could not be resolved)
    STATUS=TOOMANY   -> exit 2   (more distinct key values than the safety cap)
    STATUS=UNSAVED   -> exit 2   (backup requested but workbook never saved)
    (combridge's own non-zero codes — 3/4/5/connect failure — pass through verbatim)

Exit 2 is used for every "precondition not met / nothing done" case so the
caller can distinguish a *guarded refusal* (2) from a *crash* (3/4/5) from
*success* (0).

================================================================================
HOW COMBRIDGE IS LOCATED (portability rule)
================================================================================
combridge is NOT bundled in this project repo. Apps are authored here and
DEPLOYED into a ScripTree install that ships ``lib/combridge/combridge.exe``.
This shim finds it by walking UP the directory tree from its own location
looking for ``lib/combridge/combridge.exe`` — a relative discovery, so no
absolute path is ever baked in and the app works at whatever depth it lands.
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

# The first-line marker the generated .csx prints. Must match the literal used
# in sheet_segregator.csx.template. Kept deliberately unlikely to collide with
# real worksheet content.
SENTINEL = "__XLSEG__"

# Relative location of combridge inside a deployed ScripTree install.

# combridge plugin name for Excel automation.
PLUGIN = "excel"



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe to drop inside a C# double-quoted string.

    The generated ``.csx`` embeds user-supplied text (the key-column specifier
    and the sheet-name prefix) directly inside ``"..."`` literals. Without
    escaping, a backslash or double-quote in the user's input would either break
    compilation or — worse — let crafted input alter the script. We escape the
    four characters that are special inside a C# regular (non-verbatim) string
    literal plus the whitespace controls, in an order that is safe (backslash
    first so we don't double-escape the escapes we add afterwards).
    """
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, key_column: str, has_header: bool,
               output_mode: str, sheet_name_prefix: str, backup_first: bool,
               max_groups: int, output_format: str) -> str:
    """Fill the ``.csx`` template with the validated form values.

    String values go through :func:`csharp_literal`. Booleans are emitted as the
    C# keywords ``true``/``false``. The integer cap is emitted as a bare numeric
    literal (no quotes). Every ``__TOKEN__`` in the template must be replaced
    here — a leftover token would be a compile error, which is why the offline
    render check in the project tests greps the output for ``__[A-Z_]+__``.
    """
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    return (
        template.replace("__KEY_COLUMN__", csharp_literal(key_column))
        .replace("__HAS_HEADER__", cs_bool(has_header))
        .replace("__OUTPUT_MODE__", csharp_literal(output_mode))
        .replace("__SHEET_NAME_PREFIX__", csharp_literal(sheet_name_prefix))
        .replace("__BACKUP_FIRST__", cs_bool(backup_first))
        .replace("__MAX_GROUPS__", str(max_groups))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    """Parse argv, generate the ``.csx``, run combridge, translate the result.

    Returns the process exit code ScripTree should see (see the status→exit
    contract in the module docstring).
    """
    parser = argparse.ArgumentParser(
        description="Split the active Excel sheet into one sheet per distinct "
                    "value in a chosen key column.")
    parser.add_argument(
        "--key-column", dest="key_column", required=True,
        help="Header name, column letter (A, B, ...), or 1-based position "
             "within the used range identifying the column to split on.")
    parser.add_argument(
        "--has-header", dest="has_header", action="store_true",
        help="First row of the used range is a header row (copied to each "
             "output sheet and searchable by name for --key-column).")
    parser.add_argument(
        "--output-mode", dest="output_mode", default="new_workbook",
        choices=["new_workbook", "new_sheets"],
        help="Where to write the per-value sheets.")
    parser.add_argument(
        "--sheet-name-prefix", dest="sheet_name_prefix", default="",
        help="Optional prefix prepended to every created sheet name.")
    parser.add_argument(
        "--backup-first", dest="backup_first", action="store_true",
        help="(new_sheets mode only) Save a <name>_backup copy of the workbook "
             "before adding sheets.")
    parser.add_argument(
        "--max-groups", dest="max_groups", type=int, default=50,
        help="Safety cap: refuse if the key column has more distinct values "
             "than this (guards against picking a near-unique column).")
    parser.add_argument(
        "--output-format", dest="output_format", default="markdown",
        choices=["markdown", "text"])
    args = parser.parse_args()

    # --- defensive clamping ------------------------------------------------
    # The form's spinbox already enforces min=1, but a hand-edited config
    # sidecar or a future round-trip could feed an out-of-range value, so we
    # clamp here too (cheap, and keeps the .csx logic simple).
    max_groups = max(1, args.max_groups)

    key_column = args.key_column.strip()
    if not key_column:
        print("ERROR: --key-column must not be empty.", file=sys.stderr)
        return 2

    here = Path(__file__).resolve().parent
    template_path = here / "sheet_segregator.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        key_column=key_column,
        has_header=args.has_header,
        output_mode=args.output_mode,
        sheet_name_prefix=args.sheet_name_prefix,
        backup_first=args.backup_first,
        max_groups=max_groups,
        output_format=args.output_format,
    )

    # Write the generated script to a temp .csx, run it, and always clean up.
    # delete=False + manual unlink in finally is required on Windows because the
    # child process must be able to open the file while we hold no lock on it.
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="sheet_segregator_",
        delete=False, encoding="utf-8")
    try:
        tmp.write(csx)
        tmp.close()
        proc = subprocess.run(
            ["combridge.exe", PLUGIN, "run-script", tmp.name, "-"],
            capture_output=True, text=True,
            encoding="utf-8", errors="replace",
        )
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass

    if proc.stderr:
        sys.stderr.write(proc.stderr)

    # combridge failed before/while running the script (compile=3, throw=4,
    # host=5, or could not connect to Excel): surface its output verbatim and
    # propagate its exit code unchanged.
    if proc.returncode != 0:
        sys.stdout.write(proc.stdout)
        return proc.returncode

    # Clean run: parse the sentinel first line, print the rest as the report.
    lines = proc.stdout.splitlines()
    status = "OK"
    body_start = 0
    if lines and lines[0].startswith(SENTINEL):
        fields = lines[0].split()
        for f in fields[1:]:
            if f.startswith("STATUS="):
                status = f[len("STATUS="):]
        body_start = 1

    body = "\n".join(lines[body_start:])
    if body:
        print(body)

    # Any guarded refusal maps to exit 2; only a real success is 0.
    return 0 if status == "OK" else 2


if __name__ == "__main__":
    raise SystemExit(main())
