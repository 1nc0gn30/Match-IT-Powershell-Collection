param(
    # --- Files & sheet ---
    [string]$ExcelPath,
    [string]$SheetName,
    [string]$CsvPath,

    # --- Primary match config (Column D logic) ---
    [string]$ExcelPrimaryColumn,        # Excel column letter for primary key (ex: "D")
    [string]$CsvPrimaryMatchColumn,     # CSV column name to match on
    [string]$CsvPrimaryReturnColumnG,   # CSV column to return value for Excel Column G
    [string]$CsvPrimaryReturnColumnH,   # CSV column to return value for Excel Column H

    # --- Secondary match config (Column B logic) ---
    [string]$ExcelSecondaryColumn,       # Excel column letter for secondary key (ex: "B")
    [string]$CsvSecondaryMatchColumn,    # CSV column name to match on
    [string]$CsvSecondaryReturnColumnG,  # CSV column to return value for Excel Column G
    [string]$CsvSecondaryReturnColumnH,  # CSV column to return value for Excel Column H

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

# --- Validate CSV columns ---
foreach ($col in @(
    $CsvPrimaryMatchColumn,
    $CsvPrimaryReturnColumnG, $CsvPrimaryReturnColumnH,
    $CsvSecondaryMatchColumn,
    $CsvSecondaryReturnColumnG, $CsvSecondaryReturnColumnH
)) {
    if (-not ($csv[0].PSObject.Properties.Name -contains $col)) {
        throw "CSV column '$col' not found in CSV headers."
    }
}

# --- Build lookup tables ---
$primaryLookupG = @{}
$primaryLookupH = @{}
foreach ($row in $csv) {
    $k = $row.$CsvPrimaryMatchColumn
    if ($k) {
        $primaryLookupG[$k] = $row.$CsvPrimaryReturnColumnG
        $primaryLookupH[$k] = $row.$CsvPrimaryReturnColumnH
    }
}

$secondaryLookupG = @{}
$secondaryLookupH = @{}
foreach ($row in $csv) {
    $k = $row.$CsvSecondaryMatchColumn
    if ($k) {
        $secondaryLookupG[$k] = $row.$CsvSecondaryReturnColumnG
        $secondaryLookupH[$k] = $row.$CsvSecondaryReturnColumnH
    }
}

Write-Host "Primary rows loaded: $($primaryLookupG.Count)"
Write-Host "Secondary rows loaded: $($secondaryLookupG.Count)"

# --- Column indexes for Excel ---
$colG = Get-ColumnIndex "G"
$colH = Get-ColumnIndex "H"
$primaryIdx   = Get-ColumnIndex $ExcelPrimaryColumn
$secondaryIdx = Get-ColumnIndex $ExcelSecondaryColumn

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

        # Read the current row values in G & H
        $g = $sheet.Cells.Item($i, $colG).Text
        $h = $sheet.Cells.Item($i, $colH).Text

        $fillG = [string]::IsNullOrWhiteSpace($g) -or $g -eq "#N/A"
        $fillH = [string]::IsNullOrWhiteSpace($h) -or $h -eq "#N/A"

        if (-not ($fillG -or $fillH)) { continue }

        # Try primary match
        $primaryKey = $sheet.Cells.Item($i, $primaryIdx).Text
        $matchFound = $false
        $didPrimary = $false
        $keyUsed = $null

        if (-not [string]::IsNullOrWhiteSpace($primaryKey)) {
            if ($primaryLookupG.ContainsKey($primaryKey)) {
                $matchFound = $true
                $didPrimary = $true
                $keyUsed = $primaryKey
            }
        }

        # Try secondary only if primary failed
        if (-not $matchFound) {
            $secondaryKey = $sheet.Cells.Item($i, $secondaryIdx).Text
            if (-not [string]::IsNullOrWhiteSpace($secondaryKey)) {
                if ($secondaryLookupG.ContainsKey($secondaryKey)) {
                    $matchFound = $true
                    $didPrimary = $false
                    $keyUsed = $secondaryKey
                }
            }
        }

        if (-not $matchFound) { continue }

        # Fill G and/or H with correct values
        if ($didPrimary) {
            if ($fillG) { 
                $sheet.Cells.Item($i, $colG).Value2 = $primaryLookupG[$keyUsed]
                $updates++
            }
            if ($fillH) { 
                $sheet.Cells.Item($i, $colH).Value2 = $primaryLookupH[$keyUsed]
                $updates++
            }
        }
        else {
            if ($fillG) { 
                $sheet.Cells.Item($i, $colG).Value2 = $secondaryLookupG[$keyUsed]
                $updates++
            }
            if ($fillH) { 
                $sheet.Cells.Item($i, $colH).Value2 = $secondaryLookupH[$keyUsed]
                $updates++
            }
        }
    }

    $wb.Save()
    Write-Host "Total updated cells: $updates"
}
finally {
    if ($wb) { $wb.Close() | Out-Null }
    if ($excel) { $excel.Quit() | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Done."
