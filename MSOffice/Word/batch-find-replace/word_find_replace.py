#!/usr/bin/env python3
"""Batch Find & Replace (Word) — ScripTree shim (Strategy A).

ScripTree launches this script (executable = ``python``) with the form
values as argv.  combridge's ``run-script`` has no argv channel — a
``.csx`` only sees the plugin globals plus environment — so we BAKE the
form values into a generated ``.csx`` rendered from
``word_find_replace.csx.template``, then hand that to::

    combridge.exe word run-script <temp.csx> -

combridge's ScriptHost ignores the script's ``return`` value (it exits 0
on any clean run), so the ``.csx`` emits a first-line sentinel and THIS
shim owns the process exit code that ScripTree sees:

    __WORDREP__ STATUS=NODOC          -> exit 2 (no document open)
    __WORDREP__ STATUS=UNSAVED        -> exit 2 (copy mode needs a saved doc)
    __WORDREP__ STATUS=EMPTY_FIND     -> exit 2 (nothing to search for)
    __WORDREP__ STATUS=OK replaced=.. -> exit 0

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

SENTINEL = "__WORDREP__"



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe inside a C# double-quoted string."""
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, find_text: str, replace_text: str,
               match_case: bool, whole_word: bool, use_wildcards: bool,
               work_on_copy: bool, output_format: str) -> str:
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    return (
        template.replace("__FIND_TEXT__", csharp_literal(find_text))
        .replace("__REPLACE_TEXT__", csharp_literal(replace_text))
        .replace("__MATCH_CASE__", cs_bool(match_case))
        .replace("__WHOLE_WORD__", cs_bool(whole_word))
        .replace("__USE_WILDCARDS__", cs_bool(use_wildcards))
        .replace("__WORK_ON_COPY__", cs_bool(work_on_copy))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Find & replace text throughout the running Word document.")
    parser.add_argument("--find", dest="find_text", required=True)
    parser.add_argument("--replace", dest="replace_text", default="")
    parser.add_argument("--match-case", dest="match_case", action="store_true")
    parser.add_argument("--whole-word", dest="whole_word", action="store_true")
    parser.add_argument("--use-wildcards", dest="use_wildcards",
                        action="store_true")
    parser.add_argument("--work-on-copy", dest="work_on_copy",
                        action="store_true")
    parser.add_argument("--output-format", default="markdown",
                        choices=["markdown", "text"])
    args = parser.parse_args()

    if not args.find_text:
        print("ERROR: --find must not be empty.", file=sys.stderr)
        return 2

    here = Path(__file__).resolve().parent
    template_path = here / "word_find_replace.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        find_text=args.find_text,
        replace_text=args.replace_text,
        match_case=args.match_case,
        whole_word=args.whole_word,
        use_wildcards=args.use_wildcards,
        work_on_copy=args.work_on_copy,
        output_format=args.output_format,
    )

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="word_find_replace_",
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
        fields = lines[0].split()
        for f in fields[1:]:
            if f.startswith("STATUS="):
                status = f[len("STATUS="):]
        body_start = 1

    body = "\n".join(lines[body_start:])
    if body:
        print(body)

    # Precondition failures (no doc / unsaved in copy-mode / empty find) -> 2.
    return 2 if status in ("NODOC", "UNSAVED", "EMPTY_FIND") else 0


if __name__ == "__main__":
    raise SystemExit(main())
