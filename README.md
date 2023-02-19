# Python XLS-to-CSV Processing Script

This script opens an `employee.xls` file at a provided path with the format:

| employeeid | expense1 | expense2 | expense3 | totalexpense |
|------------|----------|----------|----------|--------------|
| 1          | 1        | 2        | 3        |              |
| 2          | 10.10    | 31       | 4        |              |
| ...        |          |          |          |              |

It then sums the columns from `expense1`, `expense2`, and `expense3` and writes
a new verson of this file, `employee.csv`, with the result of the summed columns
in the `totalexpense` column.

## Setup

This project is an unpackaged Python script. You will need Python 3.8 and pip to
run it - please make sure it is installed and in your PATH before continuing.

Download or clone this repository:
```
git clone https://github.com/bclindner/python-xls-csv-convert
cd python-xls-csv-convert
```

A virtual environment would be recommended, to avoid polluting your system's
Python installation.

Linux instructions:

```sh
python -m venv venv
. venv/bin/activate
```

Windows instructions:
```ps1
python -m venv venv
venv/Scripts/activate.ps1
```

Once your virtual environment is initialized, install the requirements:

```
pip install -r requirements.txt
```

You can now launch the script - run this command with the directory of the XLS
file as an argument:

```ps1
python run.py C:\temp\incoming\employee.xls
```

If everything was set up correctly, the program should have created a new file,
`output.csv`, in the same folder as the `employee.xls` file.

The script is configured to fail on the following cases, for safety:
* A path was not specified when running
* The path specified is not a file or does not exist
* The output.csv file already exists
