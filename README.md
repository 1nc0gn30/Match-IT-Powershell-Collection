# Excel CSV Matcher

A lightweight, configurable PowerShell utility that matches values from an Excel sheet against a CSV file, then writes returned CSV values back into the Excel workbook. This tool is ideal for data cleaning, reconciliation, and fast lookup operations without using Excel formulas or external modules.

## Features
- Map Excel column values to CSV values using any matching column.
- Return any CSV column into any Excel column.
- Supports Excel column letters (A, B, AA...) or numeric indices.
- Pure PowerShell — no external dependencies.
- Fully parameterized for automation and reuse.

## Usage

```powershell
.\match.ps1 `
    -ExcelPath "C:\Data\Employees.xlsx" `
    -SheetName "Sheet1" `
    -CsvPath "C:\Data\Directory.csv" `
    -ExcelMatchColumn "A" `
    -CsvMatchColumn "EmpID" `
    -CsvReturnColumn "ManagerName" `
    -ExcelOutputColumn "B" `
    -HeaderRow 1
```

## How It Works
1. Loads the CSV and validates required columns.
2. Builds a fast lookup hashtable based on your matching column.
3. Opens Excel through COM automation.
4. Iterates each row and writes matched values to your chosen output column.
5. Saves and closes Excel safely with full COM cleanup.

## Ideal For
- SOX analysts
- Data reconciliation tasks
- Mapping IDs to names or attributes
- Automating Excel updates without formulas
- Bulk lookups

## Requirements
- Windows
- PowerShell 5+
- Excel installed (COM automation)

## License
MIT License
