#!/usr/bin/env python3
"""Formula-to-Value Freezer (Excel) — ScripTree shim (Strategy A).

================================================================================
WHAT THIS PROGRAM DOES (the logic — read this to understand/rewrite the tool)
================================================================================
The user has an Excel workbook full of *formulas* (``=SUM(...)``, cross-sheet
references, ``=A2*B2`` columns, and so on). They want to "lock down" the current
calculated state — replace every formula with the static value it currently
produces — so the numbers can never silently change again (e.g. a month-end
snapshot, a figure handed to an auditor, or a file sent outside the company
where the source data may be missing).

The whole operation is **safe by default**: with "Work on a copy" left on (the
default), the tool saves a sibling ``<name>_Frozen.xlsx`` on disk, freezes the
formulas *there*, saves and closes that copy, and leaves the user's original
workbook **open and completely untouched** in their running Excel. The user's
live formulas survive in the instance they are looking at; only the new file on
disk is frozen.

If the user explicitly turns "Work on a copy" OFF, the tool freezes the formulas
in the *open* workbook in memory and leaves it **open and UNSAVED** for the user
to review (Ctrl+Z still undoes it). The tool never silently overwrites the
original file on disk in either mode.

================================================================================
WHY THIS IS A SHIM (Strategy A — the project-wide integration pattern)
================================================================================
ScripTree is a GUI form that runs a command line. It launches this script with
``executable = python`` and passes the form field values as argv (see the
``argument_template`` in ``formula-freezer.scriptree``). The actual Excel work
must happen inside combridge (which owns the COM connection to the running
Excel), and combridge's ``run-script`` subcommand executes a C# Roslyn script
(``.csx``).

combridge's ``run-script`` has **no argv channel** — a ``.csx`` can only see the
plugin globals (``xlApp``/``xlBook``/``xlSheet``) plus environment variables. So
this shim BAKES the form values into a generated ``.csx`` (rendered from
``freeze_formulas.csx.template`` by substituting ``__PLACEHOLDER__`` tokens),
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
     line:  ``__XLFREEZE__ STATUS=<code> [key=value ...]``  followed by the
     human-readable report.
  2. THIS shim parses that sentinel, strips it from the output, prints the
     report, and **translates the status into the process exit code** that
     ScripTree sees.

Status -> exit-code contract (see ``main`` and the ``.csx`` template):

    STATUS=OK         -> exit 0   (freeze succeeded; report follows)
    STATUS=NOBOOK     -> exit 2   (no workbook open)
    STATUS=PROTECTED  -> exit 2   (a target sheet is protected; can't write)
    STATUS=UNSAVED    -> exit 2   (copy mode requested but workbook never saved)
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
Running the shim from THIS repo (which has no bundled combridge) returns the
"could not locate" error with exit 1 — that is correct; the app is deploy-only.
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

# The first-line marker the generated .csx prints. Must match the literal used
# in freeze_formulas.csx.template. Kept deliberately unlikely to collide with
# real worksheet content.
SENTINEL = "__XLFREEZE__"

# Relative location of combridge inside a deployed ScripTree install.

# combridge plugin name for Excel automation.
PLUGIN = "excel"

# Statuses that mean "precondition not met / nothing done" -> non-zero exit (2).
PRECONDITION_FAILS = ("NOBOOK", "PROTECTED", "UNSAVED")



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe to drop inside a C# double-quoted string.

    The generated ``.csx`` embeds the (enum) form values directly inside
    ``"..."`` literals. Without escaping, a backslash or double-quote would
    either break compilation or — worse — let crafted input alter the script. We
    escape the four characters that are special inside a C# regular
    (non-verbatim) string literal plus the whitespace controls, in an order that
    is safe (backslash first so we don't double-escape the escapes we add
    afterwards). The current form has no free-text params, but this keeps the
    renderer robust if one is added later and matches the catalog's shims.
    """
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, scope: str, work_on_copy: bool,
               output_format: str) -> str:
    """Fill the ``.csx`` template with the validated form values.

    String/enum values go through :func:`csharp_literal`. Booleans are emitted
    as the C# keywords ``true``/``false``. Every ``__TOKEN__`` in the template
    must be replaced here — a leftover token would be a compile error, which is
    why the offline render check greps the output for ``__[A-Z_]+__`` and
    expects only the ``__XLFREEZE__`` sentinel to survive.
    """
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    return (
        template.replace("__SCOPE__", csharp_literal(scope))
        .replace("__WORK_ON_COPY__", cs_bool(work_on_copy))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    """Parse argv, generate the ``.csx``, run combridge, translate the result.

    Returns the process exit code ScripTree should see (see the status->exit
    contract in the module docstring).
    """
    parser = argparse.ArgumentParser(
        description="Convert every formula in the open Excel workbook into its "
                    "static value, on a copy by default.")
    parser.add_argument(
        "--scope", dest="scope", default="all_sheets",
        choices=["all_sheets", "active_sheet"],
        help="Freeze formulas on every worksheet, or only the active sheet.")
    parser.add_argument(
        "--work-on-copy", dest="work_on_copy", action="store_true",
        help="SAFETY GUARD: freeze a saved <name>_Frozen.xlsx copy and leave "
             "the original open and untouched. If omitted, the open workbook is "
             "frozen in memory and left open and UNSAVED.")
    parser.add_argument(
        "--output-format", dest="output_format", default="markdown",
        choices=["markdown", "text"])
    args = parser.parse_args()

    here = Path(__file__).resolve().parent
    template_path = here / "freeze_formulas.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        scope=args.scope,
        work_on_copy=args.work_on_copy,
        output_format=args.output_format,
    )

    # Write the generated script to a temp .csx, run it, and always clean up.
    # delete=False + manual unlink in finally is required on Windows because the
    # child process must be able to open the file while we hold no lock on it.
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="freeze_formulas_",
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
        for f in lines[0].split()[1:]:
            if f.startswith("STATUS="):
                status = f[len("STATUS="):]
        body_start = 1

    body = "\n".join(lines[body_start:])
    if body:
        print(body)

    # Any guarded refusal maps to exit 2; only a real success is 0.
    return 2 if status in PRECONDITION_FAILS else 0


if __name__ == "__main__":
    raise SystemExit(main())
