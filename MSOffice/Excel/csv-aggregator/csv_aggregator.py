#!/usr/bin/env python3
"""Multi-CSV Sheet Aggregator (Excel) — ScripTree shim (Strategy A).

ScripTree launches this script (executable = ``python``) with the form
values as argv.  combridge's ``run-script`` has no argv channel — a
``.csx`` only sees the plugin globals plus environment — so we BAKE the
form values into a generated ``.csx`` rendered from
``csv_aggregator.csx.template``, then hand that to::

    combridge.exe excel run-script <temp.csx> -

combridge's ScriptHost ignores the script's ``return`` value (it exits 0
on any clean run), so the ``.csx`` emits a first-line sentinel and THIS
shim owns the process exit code that ScripTree sees:

    __CSVAGG__ STATUS=NOWB      -> exit 2 (no workbook open)
    __CSVAGG__ STATUS=UNSAVED   -> exit 2 (backup requested but workbook never saved)
    __CSVAGG__ STATUS=NOFILES   -> exit 2 (no .csv/.txt files in the chosen folder)
    __CSVAGG__ STATUS=OK ...    -> exit 0

The tool IMPORTS each CSV/TXT in a folder as a new sheet in the ALREADY-OPEN
active workbook (a mutation), so it offers a backup-first guard
(``SaveCopyAs`` to ``<name>_Backup<ext>``) and otherwise leaves the workbook
open-and-unsaved for the user to review and save.

combridge is located by walking up from this file looking for
``lib/combridge/combridge.exe`` — a relative discovery so the catalog
stays portable (no absolute path baked in, per the project's path rule).
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

SENTINEL = "__CSVAGG__"



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe inside a C# double-quoted string."""
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, source_folder: str, delimiter: str,
               backup_first: bool, output_format: str) -> str:
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    return (
        template.replace("__SOURCE_FOLDER__", csharp_literal(source_folder))
        .replace("__DELIMITER__", csharp_literal(delimiter))
        .replace("__BACKUP_FIRST__", cs_bool(backup_first))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Import every CSV/TXT in a folder as a new sheet in the "
                    "running Excel workbook.")
    parser.add_argument("--source-folder", dest="source_folder", required=True)
    parser.add_argument("--delimiter", default="auto",
                        choices=["auto", "comma", "tab", "semicolon"])
    parser.add_argument("--backup-first", dest="backup_first",
                        action="store_true")
    parser.add_argument("--output-format", default="markdown",
                        choices=["markdown", "text"])
    args = parser.parse_args()

    if not args.source_folder.strip():
        print("ERROR: --source-folder must not be empty.", file=sys.stderr)
        return 2

    here = Path(__file__).resolve().parent
    template_path = here / "csv_aggregator.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        source_folder=args.source_folder,
        delimiter=args.delimiter,
        backup_first=args.backup_first,
        output_format=args.output_format,
    )

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="csv_aggregator_",
        delete=False, encoding="utf-8")
    try:
        tmp.write(csx)
        tmp.close()
        proc = subprocess.run(
            ["combridge.exe", "excel", "run-script", tmp.name, "-"],
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

    # combridge failed before/while running the script (compile/host/connect
    # error, or no Excel session): surface its output verbatim and its code.
    if proc.returncode != 0:
        sys.stdout.write(proc.stdout)
        return proc.returncode

    # Parse the sentinel first line written by the .csx.
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

    # Precondition failures (no workbook / unsaved-in-backup-mode / no files) -> 2.
    return 2 if status in ("NOWB", "UNSAVED", "NOFILES") else 0


if __name__ == "__main__":
    raise SystemExit(main())
