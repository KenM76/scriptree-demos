#!/usr/bin/env python3
"""Generate examples/sample_list.xlsx for the Draft Generator demo.

Writes a tiny three-row mailing list with headers Email, Name, Company using
openpyxl. Reproducible: re-running overwrites the file with the same content.
ASCII-only output (the Bash tool encodes console output as cp1252, so any
non-ASCII print() would raise UnicodeEncodeError).

Run:
    python make_example.py
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).resolve().parent
OUT = HERE / "sample_list.xlsx"

# Header row + three rows of obviously-fake data. The columns line up with the
# tokens used in the examples/README.md walkthrough:
#   {Name}    -> Name column
#   {Company} -> Company column
# The Email column is the recipient column (default header name "Email").
HEADERS = ["Email", "Name", "Company"]
ROWS = [
    ["ada@example.com",    "Ada Lovelace",    "Analytical Engines Ltd"],
    ["alan@example.com",   "Alan Turing",     "Bletchley Works"],
    ["grace@example.com",  "Grace Hopper",    "Compiler Co"],
]


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Recipients"
    ws.append(HEADERS)
    for row in ROWS:
        ws.append(row)
    wb.save(str(OUT))
    print("Wrote", OUT)
    print("Rows (excluding header):", len(ROWS))


if __name__ == "__main__":
    main()
