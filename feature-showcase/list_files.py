#!/usr/bin/env python3
"""
A choices_provider stub for the feature-showcase demo.

ScripTree runs this script when the form opens (or when an upstream
field changes) and reads one JSON document from stdout. The runtime
contract is documented in:
    D:\\Dev\\ScripTree\\help\\LLM\\dynamic_providers.md

Contract recap:
  - ScripTree writes a single JSON object to our stdin describing any
    upstream values, e.g.  {"depends_on": {"target_dir": "C:/x"},
                            "param_id": "preview_file"}
  - We print ONE JSON document to stdout. For an enum/multiselect this
    is {"choices": [...], "choice_labels": [...], "default": "..."}
  - Exit 0 on success, non-zero on failure.
  - The provider is NOT a shell command; ScripTree spawns argv
    directly. No quoting, no escaping needed.

This stub lists the files (not directories) in the directory chosen
by the upstream 'target_dir' field, sorted by name, capped at 40
entries. If no directory is supplied (or it doesn't exist) we fall
back to listing the directory that contains this script.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path


def load_stdin_payload() -> dict:
    """Read a JSON object from stdin if present; tolerate empty input."""
    raw = sys.stdin.read().strip()
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # Garbled payload from the host — fall back to no-deps mode.
        return {}


def pick_target_directory(payload: dict) -> Path:
    deps = payload.get("depends_on") or {}
    candidate = (deps.get("target_dir") or "").strip()
    if candidate:
        p = Path(candidate).expanduser()
        if p.is_dir():
            return p
    # Fallback: the directory this script lives in.
    return Path(__file__).resolve().parent


def list_files(directory: Path, cap: int = 40) -> list[Path]:
    try:
        entries = [e for e in directory.iterdir() if e.is_file()]
    except (PermissionError, OSError):
        return []
    entries.sort(key=lambda p: p.name.lower())
    return entries[:cap]


def main() -> int:
    payload = load_stdin_payload()
    target = pick_target_directory(payload)
    files = list_files(target)

    choices: list[str] = [str(p) for p in files]
    labels: list[str] = [
        f"{p.name}  ({p.stat().st_size:,} bytes)" for p in files
    ]

    out = {
        "choices": choices,
        "choice_labels": labels,
        # Pre-select the first entry if the upstream field is empty so
        # the form has a runnable default at first paint.
        "default": choices[0] if choices else "",
    }
    sys.stdout.write(json.dumps(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
