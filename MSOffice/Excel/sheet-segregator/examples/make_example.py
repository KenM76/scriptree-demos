#!/usr/bin/env python3
"""Generate the sample sales export for the Column-Based Sheet Segregator.

Run:  python make_example.py
Produces:  sample_sales.xlsx  (next to this script)

WHAT THE SAMPLE CONTAINS (and why)
----------------------------------
A single flat table on one sheet ("Sales") with a header row and 12 data
rows. The columns are:

    Date | Region | Product | Units | Revenue

The **Region** column has exactly **four** distinct values — East, West,
North, South — deliberately interleaved (not grouped) so you can see the
tool gather scattered rows. Splitting on Region therefore yields four
output sheets (East/West/North/South), each carrying a copy of the header
row plus only its own region's rows.

The row counts per region are uneven on purpose (East 4, West 3, North 3,
South 2) so the per-sheet totals in the report are easy to eyeball.
"""
from pathlib import Path

from openpyxl import Workbook

ROWS = [
    # Date,        Region, Product,   Units, Revenue
    ("2026-01-03", "East",  "Widget",  120,   2400),
    ("2026-01-03", "West",  "Gadget",   80,   3200),
    ("2026-01-04", "North", "Widget",   45,    900),
    ("2026-01-05", "East",  "Gizmo",    60,   1800),
    ("2026-01-06", "South", "Gadget",   30,   1200),
    ("2026-01-07", "West",  "Widget",   95,   1900),
    ("2026-01-08", "North", "Gizmo",    50,   1500),
    ("2026-01-09", "East",  "Gadget",   70,   2800),
    ("2026-01-10", "South", "Widget",   40,    800),
    ("2026-01-11", "West",  "Gizmo",    25,    750),
    ("2026-01-12", "North", "Gadget",   55,   2200),
    ("2026-01-13", "East",  "Widget",  110,   2200),
]


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["Date", "Region", "Product", "Units", "Revenue"])
    for r in ROWS:
        ws.append(list(r))

    out = Path(__file__).resolve().parent / "sample_sales.xlsx"
    wb.save(str(out))
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
