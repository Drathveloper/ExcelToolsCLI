# ExcelToolsCLI

Simple command-line tool for working with both Excel files (.xlsx) and CSV files (.csv).

It can perform three operations:
  - Given an Excel workbook and a CSV file, append in a given worksheet all data readed from CSV file.
  - Given an Excel workbook and a CSV file, merge in a given worksheet all data readed from CSV file, ignoring duplicated rows.
  - Given an Excel workbook and a path for a CSV file, create a CSV file containing all rows from the given worksheet.

As a command-line tool, we need to provide all the parameters and valid information. These are all the supported parameters:

Mandatory parameters:
  - -f: The Excel file path
  - -s: The sheet name
  - -m: The operation mode. Actually it supports three: merge, append and parse
  - -o: The CSV file path. In merge and append mode read from this path, and in parse mode write the csv to this path.

Optional parameters:
  - -df: Specifies the way of formatting dates. It's useful to convert strings from CSV to the custom excel date format. Supports every date format. Default: "dd/MM/yyyy"
  - -enc: Set the encoding for reading CSV files. Default: "UTF-8"
  - -i: This flag switches the way it handles formulas. If this flag is specified, it wont parse formulas and treat rows with only formula values as empty rows

