param(
    # --- Files & sheet ---
    [string]$ExcelPath          = "C:\Data\Employees.xlsx",
    [string]$SheetName          = "Sheet1",
    [string]$CsvPath            = "C:\Data\Directory.csv",

    # --- Matching config ---
    # Excel column to read from (letter like 'A' or number like '1')
    [string]$ExcelMatchColumn   = "A",

    # CSV column header to compare against
    [string]$CsvMatchColumn     = "EmpID",

    # CSV column header whose value you want to return to Excel
    [string]$CsvReturnColumn    = "ManagerName",

    # Excel column to write the returned value into (letter or number)
    [string]$ExcelOutputColumn  = "B",

    # Row that contains headers in Excel (data starts after this row)
    [int]$HeaderRow             = 1
)

function Get-ColumnIndex {
    param(
        [Parameter(Mandatory)]
        [string]$Column
    )

    # If already numeric, just return it
    if ($Column -match '^\d+$') {
        return [int]$Column
    }

    # Convert Excel-style letters (A, B, AA, AB, etc.) to 1-based index
    $col = $Column.ToUpper()
    $index = 0
    foreach ($ch in $col.ToCharArray()) {
        $index = $index * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $index
}

Write-Host "Loading CSV from '$CsvPath'..."
$csv = Import-Csv -Path $CsvPath

if (-not $csv) {
    throw "CSV file '$CsvPath' returned no rows or could not be read."
}

# Validate CSV columns
if (-not ($csv[0].PSObject.Properties.Name -contains $CsvMatchColumn)) {
    throw "CSV match column '$CsvMatchColumn' not found in CSV headers."
}
if (-not ($csv[0].PSObject.Properties.Name -contains $CsvReturnColumn)) {
    throw "CSV return column '$CsvReturnColumn' not found in CSV headers."
}

Write-Host "Building lookup from CSV column '$CsvMatchColumn' -> '$CsvReturnColumn'..."

# Build lookup hashtable: key = CSV match col, value = CSV return col
$lookup = @{}
foreach ($row in $csv) {
    $key = $row.$CsvMatchColumn
    if (-not [string]::IsNullOrWhiteSpace($key)) {
        # last one wins silently; change if you want first-win
        $lookup[$key] = $row.$CsvReturnColumn
    }
}

Write-Host "Lookup entries built: $($lookup.Count)"

# Convert Excel column letters/numbers to indices
$excelMatchIndex  = Get-ColumnIndex -Column $ExcelMatchColumn
$excelOutputIndex = Get-ColumnIndex -Column $ExcelOutputColumn

Write-Host "Opening Excel '$ExcelPath', sheet '$SheetName'..."
$excel   = $null
$workbook = $null
$sheet    = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($ExcelPath)
    $sheet    = $workbook.Worksheets.Item($SheetName)

    $usedRange = $sheet.UsedRange
    $lastRow   = $usedRange.Rows.Count

    $startRow = $HeaderRow + 1
    Write-Host "Processing rows $startRow to $lastRow..."

    $matches = 0
    for ($i = $startRow; $i -le $lastRow; $i++) {
        # Read value from Excel match column
        $value = $sheet.Cells.Item($i, $excelMatchIndex).Text

        if (-not [string]::IsNullOrWhiteSpace($value) -and $lookup.ContainsKey($value)) {
            $sheet.Cells.Item($i, $excelOutputIndex).Value2 = $lookup[$value]
            $matches++
        }
    }

    Write-Host "Rows matched and updated: $matches"

    $workbook.Save()
    Write-Host "Workbook saved."
}
finally {
    if ($workbook) { $workbook.Close() | Out-Null }
    if ($excel)    { $excel.Quit()     | Out-Null }

    if ($sheet)    { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
    if ($workbook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($excel)    { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Done."
