#!/usr/bin/env python3
"""Tracked-Changes Processing Master (Word) — ScripTree shim (Strategy A).

ScripTree launches this script (executable = ``python``) with the form
values as argv. combridge's ``run-script`` has no argv channel — a ``.csx``
only sees the plugin globals plus environment — so we BAKE the form values
into a generated ``.csx`` rendered from ``process_revisions.csx.template`` by
plain token replacement, then hand that to::

    combridge.exe word run-script <temp.csx> -

combridge's ScriptHost ignores the script's ``return`` value (it exits 0 on
any clean run), so the ``.csx`` emits a first-line sentinel and THIS shim
owns the process exit code that ScripTree sees:

    __WORDREV__ STATUS=NODOC          -> exit 2 (no document open)
    __WORDREV__ STATUS=UNSAVED        -> exit 2 (copy mode needs a saved doc)
    __WORDREV__ STATUS=OK ...         -> exit 0 (incl. the revisions=0 no-op)

NOTE on the OK/no-op case: a document with ZERO tracked changes is NOT a
precondition failure. There is simply nothing to accept or reject, so the
``.csx`` emits ``STATUS=OK revisions=0`` and this shim returns 0 ("succeeded,
nothing to do"). Only a missing document, or a never-saved document in
copy-mode, are exit-2 precondition failures.

WHAT THIS TOOL IS (and how it differs from its sibling). The Corporate Style
Sanitizer (``../style-sanitizer/``) REFUSES to run on a document that has
tracked changes — bulk-editing a document under review would bury it in
revision marks. THIS tool is the inverse: it specifically TARGETS documents
that HAVE revisions and force-accepts or force-rejects ALL of them in one
pass, clearing the revision history before the document is sent out.

combridge is invoked via a bare ``combridge.exe`` call -- ScripTree's runner
prepends ``<install>/lib/combridge`` to ``PATH`` on every spawned tool, so the
OS resolves the bundled binary by name.  No discovery code needed; if a tool is
launched outside ScripTree (e.g. direct ``python tool.py`` for debugging), set
``SCRIPTREE_HOME`` and prepend ``%SCRIPTREE_HOME%/lib/combridge`` to PATH
yourself, or just launch through ScripTree's editor. If it is not found the shim exits
1 with "could not locate ..." — that is the correct behaviour when this app
is run from the source repo (which has no combridge bundle).

WHY a shim instead of a typed plugin command: see
``_meta/rags/integration/csx_has_no_argv_channel.md`` and
``shim_generates_csx.md``. The same Strategy-A pattern backs every Office app
in this catalog.
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

SENTINEL = "__WORDREV__"
# Statuses that mean "precondition not met / nothing could be done" -> exit 2.
# revisions=0 is deliberately NOT here: it is reported under STATUS=OK (a
# successful no-op), so it maps to exit 0.
PRECONDITION_FAILS = ("NODOC", "UNSAVED")



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe inside a C# double-quoted string."""
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, action: str, work_on_copy: bool,
               output_format: str) -> str:
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    return (
        template.replace("__ACTION__", csharp_literal(action))
        .replace("__WORK_ON_COPY__", cs_bool(work_on_copy))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Force-accept or force-reject ALL tracked changes in the "
                    "running Word document in one pass.")
    parser.add_argument("--action", default="accept",
                        choices=["accept", "reject"])
    parser.add_argument("--work-on-copy", dest="work_on_copy",
                        action="store_true")
    parser.add_argument("--output-format", default="markdown",
                        choices=["markdown", "text"])
    args = parser.parse_args()

    here = Path(__file__).resolve().parent
    template_path = here / "process_revisions.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        action=args.action,
        work_on_copy=args.work_on_copy,
        output_format=args.output_format,
    )

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="process_revisions_",
        delete=False, encoding="utf-8")
    try:
        tmp.write(csx)
        tmp.close()
        proc = subprocess.run(
            ["combridge.exe", "word", "run-script", tmp.name, "-"],
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
    # error, or no Word session): surface its output verbatim and its code.
    if proc.returncode != 0:
        sys.stdout.write(proc.stdout)
        return proc.returncode

    # Parse the sentinel first line written by the .csx.
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

    return 2 if status in PRECONDITION_FAILS else 0


if __name__ == "__main__":
    raise SystemExit(main())
