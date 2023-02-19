#!/usr/bin/env python3
"""
Open an XLS file, process the expenses in each file, and save the result as a
CSV.

The file should be a table with headers (employeeid, expense1, expense2, expense3, totalexpense).
Each row after the header will be processed and saved as a CSV file.

"""
import csv
from pathlib import Path
from sys import argv, exit

import xlrd

# Headers to write to the CSV file.
FIELDNAMES = ["employeeid", "expense1", "expense2", "expense3", "totalexpense"]


# get paths and do some brief validation
try:
    assert len(argv) >= 2, "No path specified"
    xls_path = Path(argv[1])
    assert xls_path.is_file(), "Specified path is not a file"
    assert xls_path.exists(), "Specified XLS file does not exist"
    csv_path = xls_path / "./output.csv"
    assert not csv_path.exists(), "output.csv already exists"
except AssertionError as exc:
    print(f"Could not run script: {exc}")
    exit(1)

# set up csv file
with open(csv_path, "w", encoding="utf-8", newline="") as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(FIELDNAMES)
    # set up workbook
    workbook = xlrd.open_workbook(xls_path)
    # XLS files have no concept of an active sheet so we will assume the first
    # sheet here
    sheet = workbook.sheet_by_index(0)
    # we use range(1, ...) here to skip the first row
    for i in range(1, sheet.nrows):
        # save employee id
        employee_id = sheet.cell_value(rowx=i, colx=0)
        # get expense columns
        expense1 = sheet.cell_value(rowx=i, colx=1)
        expense2 = sheet.cell_value(rowx=i, colx=2)
        expense3 = sheet.cell_value(rowx=i, colx=3)
        # add and create total expense
        total_expense = sum([expense1, expense2, expense3])
        # write the row
        writer.writerow([employee_id, expense1, expense2, expense3, total_expense])
