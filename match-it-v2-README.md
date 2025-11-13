# Excel CSV Conditional Matcher

A specialized PowerShell utility that updates Excel rows **only when specific columns contain `#N/A` or are empty**.  
This version is ideal for data cleanup workflows where partial Excel rows require conditional population based on a CSV file.

---

## 🚀 Features

### ✔ Conditional Row Processing
Only updates a row when **Column G or Column H**:
- is empty  
- OR contains the string `#N/A`

### ✔ Controlled Lookup Logic
When a row qualifies:
1. PowerShell reads a value from a configured Excel column.
2. It matches that value against a CSV column.
3. It writes the corresponding CSV return value back into Excel.

### ✔ Fully Parameterized
All paths, sheet names, and column references are script parameters — nothing hard-coded.

### ✔ COM-Safe Excel Automation
Includes proper:
- COM cleanup  
- Garbage collection  
- Non-visible Excel execution  

### ✔ No External Modules Required
100% pure PowerShell.

---

## 📥 Input Requirements

### **Excel File**
Must contain:
- A column to match values from (e.g., A)
- Columns **G** and **H** (used to decide if a row is updated)
- A target output column

### **CSV File**
Must contain:
- A match column (e.g., `EmpID`)
- A return column (e.g., `ManagerName`)

---

## 🔧 Parameters

| Parameter | Description |
|----------|-------------|
| `ExcelPath` | Full path to the Excel file |
| `SheetName` | Worksheet name |
| `CsvPath` | Path to CSV file |
| `ExcelMatchColumn` | Excel column to read values from |
| `CsvMatchColumn` | CSV column to match against |
| `CsvReturnColumn` | CSV column to return data from |
| `ExcelOutputColumn` | Excel column to write returned values into |
| `HeaderRow` | Row where headers stop and data begins |

All parameters are **required** (no defaults).

---

## ▶ Example Usage

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
