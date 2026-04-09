#!/usr/bin/env python3

import argparse
import os
from datetime import datetime
from openpyxl import load_workbook


def is_parent_row(value):
    if value is None:
        return False
    return str(value).strip().startswith("Tên sản phẩm:")


def is_child_row(value):
    if value is None:
        return False

    if isinstance(value, (int, float)):
        return True

    s = str(value).strip()
    if s.isdigit():
        return True

    if s.rstrip(". )").isdigit():
        return True

    return False


def fix_file(input_path, output_path):
    wb = load_workbook(input_path)
    ws = wb.active

    current_parent = None
    changed_count = 0

    for row in range(1, ws.max_row + 1):
        cellA = ws.cell(row=row, column=1)
        cellF = ws.cell(row=row, column=6)

        # Nếu là dòng cha
        if is_parent_row(cellA.value):
            current_parent = str(cellA.value).strip()
            continue

        # Nếu là dòng con -> luôn ghi đè
        if is_child_row(cellA.value):
            if current_parent:
                cellF.value = current_parent
                changed_count += 1

    wb.save(output_path)
    return changed_count


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="Input Excel file (.xlsx)")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print("File không tồn tại")
        return

    if args.overwrite:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = f"{args.input}.backup.{ts}"
        os.rename(args.input, backup)
        output_path = args.input
        print(f"Đã backup: {backup}")
    else:
        base, ext = os.path.splitext(args.input)
        output_path = f"{base}_fixed{ext}"

    changed = fix_file(args.input, output_path)

    print("Hoàn thành.")
    print("Số ô đã cập nhật:", changed)
    print("File output:", output_path)


if __name__ == "__main__":
    main()