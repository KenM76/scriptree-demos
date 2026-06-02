#!/usr/bin/env python3
"""Mailbox Age & Size Auditor — ScripTree shim (Strategy A).

ScripTree launches this script (executable = ``python``) with the form
values as argv.  combridge's ``run-script`` has no argv channel — a
``.csx`` only sees the plugin globals plus environment — so we BAKE the
form values into a generated ``.csx`` rendered from
``mailbox_auditor.csx.template``, then hand that to::

    combridge.exe outlook run-script <temp.csx> -

combridge's ScriptHost ignores the script's ``return`` value (it exits 0
on any clean run), so the ``.csx`` emits a first-line sentinel and THIS
shim owns the process exit code that ScripTree sees:

    __MAILAGE__ STATUS=NO_STORE              -> exit 2 (precondition not met)
    __MAILAGE__ STATUS=OK folders=.. ...     -> exit 0

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

SENTINEL = "__MAILAGE__"



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe inside a C# double-quoted string."""
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, scope: str, stale_years: int,
               min_folder_mb: int, output_format: str) -> str:
    return (
        template.replace("__SCOPE__", csharp_literal(scope))
        .replace("__STALE_YEARS__", str(stale_years))
        .replace("__MIN_FOLDER_MB__", str(min_folder_mb))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Audit per-folder age and size of the running Outlook's stores.")
    parser.add_argument("--scope", default="all", choices=["all", "default"])
    parser.add_argument("--stale-years", type=int, default=2)
    parser.add_argument("--min-folder-mb", type=int, default=5)
    parser.add_argument("--output-format", default="markdown",
                        choices=["markdown", "text"])
    args = parser.parse_args()

    # Clamp to non-negative — the form enforces this, but guard anyway.
    stale_years = max(0, args.stale_years)
    min_folder_mb = max(0, args.min_folder_mb)

    here = Path(__file__).resolve().parent
    template_path = here / "mailbox_auditor.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        scope=args.scope,
        stale_years=stale_years,
        min_folder_mb=min_folder_mb,
        output_format=args.output_format,
    )

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="mailbox_auditor_",
        delete=False, encoding="utf-8")
    try:
        tmp.write(csx)
        tmp.close()
        proc = subprocess.run(
            ["combridge.exe", "outlook", "run-script", tmp.name, "-"],
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
    # error, or no Outlook session): surface its output verbatim and its code.
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

    return 2 if status == "NO_STORE" else 0


if __name__ == "__main__":
    raise SystemExit(main())
