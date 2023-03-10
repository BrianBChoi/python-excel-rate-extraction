# python-excel-rate-extraction
Script that parses a vessel carrier's service contract for relevant data and appends it to an existing file.

Using Pandas to read service contract and extract relevant data in rate tables. Data is loaded into memory before systematically being appended to a new sheet in an existing Excel file, following the end user's desired format.

Future goals:
1) Use PyInstaller so end users can execute script without using a terminal or installing Python.
2) Refactor reading and writing functions to make separation of concerns clearer.
3) Add more parsing functions for service contracts formatted differently.
