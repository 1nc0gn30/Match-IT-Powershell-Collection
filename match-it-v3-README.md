# Match-It: Dual-Stage Excel G/H Filler

This script performs **conditional row updates** in Excel using values from a CSV file.  
It is designed for cases where **Excel Columns G and H must be filled ONLY when they are empty or contain `#N/A`**, using a **two-level matching system**:

### ✅ Primary Match (Column D → CSV)  
### ✅ Secondary Match (Column B → CSV)  
### 🎯 Only updates G/H when needed  
### 🎯 Fills G, H, or both depending on which cells are empty  

This tool is perfect for reconciliation, SOX-related cleanup work, identity mapping, access list population, and automated Excel enrichment.

---

## 🔍 Core Logic Summary

For each row in Excel:

### **1. Check Column G and H**
If **Column G or Column H** is:
- empty  
- OR contains `#N/A`  

➡ Then the row **qualifies for processing**.  
If both G and H already contain valid values, the row is skipped.

---

## **2. Primary Match (Excel Column D → CSV)**

If Column **D** has a value:

- Compare that value to a specified **CSV match column**
- If found, retrieve the CSV **return column value**
- Write the retrieved value into:
  - **Column G** if G needed filling  
  - **Column H** if H needed filling  
  - **Both G and H** if both were empty/`#N/A`  

If:
- Column D is empty  
- OR the primary lookup fails  
➡ Move on to secondary matching.

---

## **3. Secondary Match (Excel Column B → CSV)**

If no primary match occurred:

- Compare Excel **Column B** to a second CSV match column
- Retrieve the associated CSV return value
- Fill it into:
  - Column G (if missing)
  - Column H (if missing)
  - Both (if both missing)

If no match is found → the row is skipped.

---

## 📥 Parameters

| Parameter | Description |
|----------|-------------|
| `ExcelPath` | Path to the Excel workbook |
| `SheetName` | The worksheet to process |
| `CsvPath` | Input CSV file |
| **Primary matching** ||
| `ExcelPrimaryColumn` | Excel column letter for primary lookup (e.g., `D`) |
| `CsvPrimaryMatchColumn` | CSV column to match against |
| `CsvPrimaryReturnColumn` | CSV column whose value will be written to Excel |
| **Secondary matching** ||
| `ExcelSecondaryColumn` | Excel column letter for fallback lookup (e.g., `B`) |
| `CsvSecondaryMatchColumn` | CSV column to match against |
| `CsvSecondaryReturnColumn` | CSV column to return if primary fails |
| `HeaderRow` | Row where headers end and data begins |

---

## ▶ Example Usage

```powershell
.\match.ps1 `
  -ExcelPath "C:\Data\Employees.xlsx" `
  -SheetName "Sheet1" `
  -CsvPath "C:\Data\Directory.csv" `
  -ExcelPrimaryColumn "D" `
  -CsvPrimaryMatchColumn "EmpID" `
  -CsvPrimaryReturnColumn "ManagerName" `
  -ExcelSecondaryColumn "B" `
  -CsvSecondaryMatchColumn "AltID" `
  -CsvSecondaryReturnColumn "ManagerName" `
  -HeaderRow 1
