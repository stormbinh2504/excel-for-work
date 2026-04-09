#!/usr/bin/env python3
"""
Propagate parent product name ("Tên sản phẩm: ...")
into column F (MST người mua) of child rows.

Parent row:
    Column A starts with "Tên sản phẩm:"

Child row:
    Column A contains numeric STT (1,2,3...)

Usage:
    python scripts/fix_excel_buyers.py input.xlsx
    python scripts/fix_excel_buyers.py input.xlsx --overwrite
"""

from __future__ import annotations

import argparse
import os
from datetime import datetime
from typing import Optional

from openpyxl import load_workbook


# -----------------------------
# Helpers
# -----------------------------

def is_parent_row(cell_value) -> bool:
    """Check if row is a parent row."""
    if cell_value is None:
        return False
    return str(cell_value).strip().startswith("Tên sản phẩm:")


def is_child_row(cell_value) -> bool:
    """Check if row is a child row (numeric STT)."""
    if cell_value is None:
        return False

    if isinstance(cell_value, (int, float)):
        return True

    s = str(cell_value).strip()
    if s.isdigit():
        return True

    if s.rstrip(". )").isdigit():
        return True

    return False


def is_empty(val) -> bool:
    return val is None or str(val).strip() == ""


# -----------------------------
# Main Processing
# -----------------------------

def fix_file(input_path: str, output_path: str) -> dict:
    wb = load_workbook(input_path)
    ws = wb.active

    current_parent_value: Optional[str] = None
    changed_count = 0

    for row in range(1, ws.max_row + 1):
        cellA = ws.cell(row=row, column=1)   # STT or Parent text
        cellF = ws.cell(row=row, column=6)   # MST người mua

        # Nếu là dòng cha
        if is_parent_row(cellA.value):
            current_parent_value = str(cellA.value).strip()
            continue

        # Nếu là dòng con
        if is_child_row(cellA.value):
            if current_parent_value and is_empty(cellF.value):
                cellF.value = current_parent_value
                changed_count += 1

    wb.save(output_path)
    return {"changed_cells": changed_count, "output": output_path}


# -----------------------------
# CLI
# -----------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Copy 'Tên sản phẩm:' from parent rows to column F (MST người mua) of child rows."
    )
    parser.add_argument("input", help="Input Excel file (.xlsx)")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite input file (creates backup)")
    args = parser.parse_args()

    input_path = args.input

    if not os.path.isfile(input_path):
        raise SystemExit(f"File not found: {input_path}")

    if args.overwrite:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = f"{input_path}.backup.{ts}"
        os.rename(input_path, backup)
        output_path = input_path
        print(f"Backup created: {backup}")
    else:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_fixed{ext}"

    print(f"Processing: {input_path}")
    result = fix_file(input_path, output_path)

    print(f"Done. Changed cells: {result['changed_cells']}")
    print(f"Output file: {result['output']}")


if __name__ == "__main__":
    main()