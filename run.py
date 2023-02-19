#!/usr/bin/env python3
"""
Open an XLSM file, process the expenses in each file, and save the result as a
CSV.

The file should be a table with headers (employeeid, expense1, expense2, expense3, totalexpense).
Each row after the header will be processed and saved as a CSV file.

"""
import csv
from pathlib import Path
from sys import argv, exit

import openpyxl

# Headers to write to the CSV file.
FIELDNAMES = ["employeeid", "expense1", "expense2", "expense3", "totalexpense"]


# get paths and do some brief validation
try:
    assert len(argv) >= 2, "No path specified"
    xlsm_path = Path(argv[1])
    assert xlsm_path.is_file(), "Specified path is not a file"
    assert xlsm_path.exists(), "Specified XLSM file does not exist"
    csv_path = (xlsm_path / ".." /  "output.csv").resolve()
    assert not csv_path.exists(), "output.csv already exists"
except AssertionError as exc:
    print(f"Could not run script: {exc}")
    exit(1)

# set up csv file
with open(csv_path, "w", encoding="utf-8", newline="") as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(FIELDNAMES)
    # set up workbook
    workbook = openpyxl.load_workbook(xlsm_path)
    # get the active sheet
    sheet = workbook.active
    # we use min_row=2 here to skip the first row
    for row in sheet.iter_rows(min_row=2):
        # save employee id
        employee_id = row[0].value
        # get expense columns
        expense1 = row[1].value
        expense2 = row[2].value
        expense3 = row[3].value
        # add and create total expense
        total_expense = sum([float(e) for e in (expense1, expense2, expense3)])
        # write the row
        writer.writerow([employee_id, expense1, expense2, expense3, total_expense])
