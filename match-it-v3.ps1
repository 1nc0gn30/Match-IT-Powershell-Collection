param(
    # --- Files & sheet ---
    [string]$ExcelPath,
    [string]$SheetName,
    [string]$CsvPath,

    # --- Primary match config (Column D logic) ---
    [string]$ExcelPrimaryColumn,       # Excel: D
    [string]$CsvPrimaryMatchColumn,    # CSV match
    [string]$CsvPrimaryReturnColumn,   # CSV return column for G/H
    # (Same return value written to G or H depending on what needs filling)

    # --- Secondary match config (Column B logic) ---
    [string]$ExcelSecondaryColumn,       # Excel: B
    [string]$CsvSecondaryMatchColumn,    # CSV match
    [string]$CsvSecondaryReturnColumn,   # CSV return column for G/H

    # Header row
    [int]$HeaderRow
)

function Get-ColumnIndex {
    param([string]$Column)
    if ($Column -match '^\d+$') { return [int]$Column }

    $col = $Column.ToUpper()
    $index = 0
    foreach ($ch in $col.ToCharArray()) {
        $index = $index * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $index
}

Write-Host "Loading CSV..."
$csv = Import-Csv -Path $CsvPath

if (-not $csv) { throw "CSV empty or unreadable." }

# Validate CSV columns
foreach ($col in @(
    $CsvPrimaryMatchColumn, $CsvPrimaryReturnColumn,
    $CsvSecondaryMatchColumn, $CsvSecondaryReturnColumn
)) {
    if (-not ($csv[0].PSObject.Properties.Name -contains $col)) {
        throw "CSV column '$col' not found."
    }
}

# Build lookup tables
$primaryLookup = @{}
foreach ($row in $csv) {
    $k = $row.$CsvPrimaryMatchColumn
    if ($k) { $primaryLookup[$k] = $row.$CsvPrimaryReturnColumn }
}

$secondaryLookup = @{}
foreach ($row in $csv) {
    $k = $row.$CsvSecondaryMatchColumn
    if ($k) { $secondaryLookup[$k] = $row.$CsvSecondaryReturnColumn }
}

Write-Host "Primary entries: $($primaryLookup.Count)"
Write-Host "Secondary entries: $($secondaryLookup.Count)"

# Column indexes
$colG = Get-ColumnIndex "G"
$colH = Get-ColumnIndex "H"
$primaryIdx   = Get-ColumnIndex $ExcelPrimaryColumn   # D
$secondaryIdx = Get-ColumnIndex $ExcelSecondaryColumn # B

Write-Host "Opening Excel..."
$excel = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($ExcelPath)
    $sheet = $wb.Worksheets.Item($SheetName)

    $lastRow = $sheet.UsedRange.Rows.Count

    $start = $HeaderRow + 1
    $updates = 0

    for ($i = $start; $i -le $lastRow; $i++) {

        # Read G & H values
        $g = $sheet.Cells.Item($i, $colG).Text
        $h = $sheet.Cells.Item($i, $colH).Text

        # Determine which columns need filling
        $fillG = [string]::IsNullOrWhiteSpace($g) -or $g -eq "#N/A"
        $fillH = [string]::IsNullOrWhiteSpace($h) -or $h -eq "#N/A"

        # If neither G nor H need filling, skip row
        if (-not ($fillG -or $fillH)) { continue }

        # ---------------------------
        # PRIMARY MATCH (Column D)
        # ---------------------------
        $primaryKey = $sheet.Cells.Item($i, $primaryIdx).Text
        $matchFound = $false
        $valueToWrite = $null

        if (-not [string]::IsNullOrWhiteSpace($primaryKey)) {
            if ($primaryLookup.ContainsKey($primaryKey)) {
                $valueToWrite = $primaryLookup[$primaryKey]
                $matchFound = $true
            }
        }

        # ---------------------------
        # SECONDARY MATCH (Column B)
        # Only run if primary failed
        # ---------------------------
        if (-not $matchFound) {
            $secondaryKey = $sheet.Cells.Item($i, $secondaryIdx).Text

            if (-not [string]::IsNullOrWhiteSpace($secondaryKey)) {
                if ($secondaryLookup.ContainsKey($secondaryKey)) {
                    $valueToWrite = $secondaryLookup[$secondaryKey]
                    $matchFound = $true
                }
            }
        }

        # If NO match found at all → skip filling
        if (-not $matchFound) { continue }

        # ---------------------------
        # Fill only G and/or H
        # ---------------------------
        if ($fillG) {
            $sheet.Cells.Item($i, $colG).Value2 = $valueToWrite
            $updates++
        }

        if ($fillH) {
            $sheet.Cells.Item($i, $colH).Value2 = $valueToWrite
            $updates++
        }

    }

    $wb.Save()
    Write-Host "Total updated cells: $updates"
}
finally {
    if ($wb) { $wb.Close() | Out-Null }
    if ($excel) { $excel.Quit() | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host "Done."
