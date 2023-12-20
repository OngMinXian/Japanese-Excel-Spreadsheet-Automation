# Japanese-ExcelSheet-Automation
The program processes an excel spreadsheet in Japanese characters by splitting the spreadsheet into multiple spreadsheets based on column 2 within the first 3 pages. This potentially saves numerous hours of work every quarter year. The script is compiled to an executable file for ease of usage for non-programmers.

## Setting up the executable
To setup the executable, Python 3.11 and its dependacies are required.
```sh
pip install -r requirements.txt
```

To create the executable, run the following commands in the directory of 'script.py'.
```sh
pyinstaller --onefile script.py
```

This creates an executable file in the 'dist' folder;

## Using the executable
The program requires a 5 page excel spreadsheet. The first 3 pages are data to be processed, with column 2 being the category to split the spreadsheet by.

