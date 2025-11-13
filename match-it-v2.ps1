param(
    # --- Files & sheet ---
    [string]$ExcelPath,
    [string]$SheetName,
    [string]$CsvPath,

    # --- Matching config ---
    [string]$ExcelMatchColumn,
    [string]$CsvMatchColumn,
    [string]$CsvReturnColumn,
    [string]$ExcelOutputColumn,

    # Header row
    [int]$HeaderRow
)

function Get-ColumnIndex {
    param([Parameter(Mandatory)][string]$Column)

    if ($Column -match '^\d+$') { return [int]$Column }

    $col = $Column.ToUpper()
    $index = 0
    foreach ($ch in $col.ToCharArray()) {
        $index = $index * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $index
}

Write-Host "Loading CSV from '$CsvPath'..."
$csv = Import-Csv -Path $CsvPath

if (-not $csv) { throw "CSV empty or unreadable." }

# Validate CSV headers
if (-not ($csv[0].PSObject.Properties.Name -contains $CsvMatchColumn)) {
    throw "CSV match column '$CsvMatchColumn' not found."
}
if (-not ($csv[0].PSObject.Properties.Name -contains $CsvReturnColumn)) {
    throw "CSV return column '$CsvReturnColumn' not found."
}

# Build lookup table
$lookup = @{}
foreach ($row in $csv) {
    $key = $row.$CsvMatchColumn
    if ($key) { $lookup[$key] = $row.$CsvReturnColumn }
}

Write-Host "Lookup entries loaded: $($lookup.Count)"

# Convert column references
$matchIdx  = Get-ColumnIndex $ExcelMatchColumn
$outputIdx = Get-ColumnIndex $ExcelOutputColumn
$colGIdx   = Get-ColumnIndex "G"
$colHIdx   = Get-ColumnIndex "H"

Write-Host "Opening Excel..."
$excel = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($ExcelPath)
    $sheet = $wb.Worksheets.Item($SheetName)

    $usedRange = $sheet.UsedRange
    $lastRow = $usedRange.Rows.Count

    $startRow = $HeaderRow + 1

    $matches = 0

    for ($i = $startRow; $i -le $lastRow; $i++) {

        # Values in G and H
        $g = $sheet.Cells.Item($i, $colGIdx).Text
        $h = $sheet.Cells.Item($i, $colHIdx).Text

        # NEW RULE:
        # Only run lookup if G OR H is empty OR "#N/A"
        $needsUpdate =
            ([string]::IsNullOrWhiteSpace($g) -or $g -eq "#N/A") -or
            ([string]::IsNullOrWhiteSpace($h) -or $h -eq "#N/A")

        if (-not $needsUpdate) {
            continue  # Skip this row entirely
        }

        # Read match value from configured column
        $value = $sheet.Cells.Item($i, $matchIdx).Text

        # Skip if match field empty (can't match against nothing)
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        # Fill the output only if match found
        if ($lookup.ContainsKey($value)) {
            $sheet.Cells.Item($i, $outputIdx).Value2 = $lookup[$value]
            $matches++
        }
    }

    $wb.Save()
    Write-Host "Updated rows: $matches"
}
finally {
    if ($wb) { $wb.Close()  | Out-Null }
    if ($excel) { $excel.Quit() | Out-Null }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Done."
