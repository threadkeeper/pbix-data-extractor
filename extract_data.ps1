<#
.SYNOPSIS
    Extracts data from a Power BI Desktop model and generates SQL INSERT statements.

.DESCRIPTION
    Connects to the local Analysis Services instance started by Power BI Desktop,
    reads all tables and their data, then generates a .sql file with INSERT statements.

.NOTES
    - Power BI Desktop must be open with the report loaded.
    - Requires ADOMD.NET (installed with Power BI Desktop or SSMS).
#>

param(
    [string]$OutputFile = "insert_data.sql",
    [int]$BatchSize = 1000
)

$ErrorActionPreference = "Stop"

# --- Find ADOMD.NET ---
$adomdPath = "C:\Program Files\Microsoft.NET\ADOMD.NET\170\Microsoft.AnalysisServices.AdomdClient.dll"
if (-not (Test-Path $adomdPath)) {
    $adomdPath = Get-ChildItem "C:\Program Files\Microsoft.NET\ADOMD.NET" -Recurse -Filter "Microsoft.AnalysisServices.AdomdClient.dll" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
    if (-not $adomdPath) {
        Write-Error "ADOMD.NET not found. Install SQL Server Analysis Services client libraries."
        exit 1
    }
}
Add-Type -Path $adomdPath

# --- Find Power BI Desktop SSAS port ---
$msmdsrv = Get-Process -Name "msmdsrv" -ErrorAction SilentlyContinue
if (-not $msmdsrv) {
    Write-Error "Power BI Desktop is not running. Open your .pbix file first."
    exit 1
}
$port = (Get-NetTCPConnection -OwningProcess $msmdsrv.Id -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1).LocalPort
if (-not $port) {
    Write-Error "Could not find SSAS listening port."
    exit 1
}
Write-Host "Found Power BI SSAS on port: $port"

# --- Connect ---
$connStr = "Data Source=localhost:$port"
$conn = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection($connStr)
$conn.Open()
Write-Host "Connected to model: $($conn.Database)"

# --- Get table list via DMV ---
$cmd = $conn.CreateCommand()
$cmd.CommandText = "SELECT [TABLE_NAME] FROM `$SYSTEM.DBSCHEMA_TABLES WHERE TABLE_TYPE = 'TABLE'"
$reader = $cmd.ExecuteReader()
$tables = @()
while ($reader.Read()) {
    $tableName = $reader[0].ToString()
    # DMV returns names with $ prefix - strip it
    if ($tableName.StartsWith('$')) {
        $tableName = $tableName.Substring(1)
    }
    # Skip internal/system tables
    if ($tableName -notlike 'LocalDateTable*' -and $tableName -notlike 'DateTableTemplate*') {
        $tables += $tableName
    }
}
$reader.Close()
Write-Host "Found $($tables.Count) tables: $($tables -join ', ')"

# --- Build TMSCHEMA column lookup (accurate types, excludes calculated/RowNumber) ---
# TMSCHEMA_COLUMNS.Type: 1 = Data, 2 = Calculated, 3 = RowNumber
# TMSCHEMA_COLUMNS.DataType: 2=String, 6=Int64, 8=Double, 9=DateTime, 10=Decimal, 11=Boolean

$tblCmd = $conn.CreateCommand()
$tblCmd.CommandText = "SELECT [ID], [Name] FROM `$SYSTEM.TMSCHEMA_TABLES"
$tblReader = $tblCmd.ExecuteReader()
$tableIdToName = @{}
$tableNameToId = @{}
while ($tblReader.Read()) {
    $tableIdToName[$tblReader[0].ToString()] = $tblReader[1].ToString()
    $tableNameToId[$tblReader[1].ToString()] = $tblReader[0].ToString()
}
$tblReader.Close()

# Load all data columns (Type=1) with their real DataType
$tmColCmd = $conn.CreateCommand()
$tmColCmd.CommandText = "SELECT [TableID], [ExplicitName], [ExplicitDataType] FROM `$SYSTEM.TMSCHEMA_COLUMNS WHERE [Type] = 1"
$tmColReader = $tmColCmd.ExecuteReader()
$tmschemaColumns = @{} # key = tableName, value = list of @{Name;DataType}
while ($tmColReader.Read()) {
    $tableId = $tmColReader[0].ToString()
    $colName = $tmColReader[1].ToString()
    $dataType = [int]$tmColReader[2]
    $tblName = $tableIdToName[$tableId]
    if (-not $tmschemaColumns.ContainsKey($tblName)) {
        $tmschemaColumns[$tblName] = [System.Collections.Generic.List[hashtable]]::new()
    }
    $tmschemaColumns[$tblName].Add(@{ Name = $colName; DataType = $dataType })
}
$tmColReader.Close()

# Count tables that have data columns vs those excluded (calculated tables)
$dataTableCount = $tmschemaColumns.Count
$excludedTableCount = $tables.Count - $dataTableCount
Write-Host "Found $dataTableCount data tables, $excludedTableCount calculated tables excluded"

# --- Get column info for a table from our TMSCHEMA cache ---
function Get-TableColumns {
    param([string]$TableName)
    if ($tmschemaColumns.ContainsKey($TableName)) {
        return $tmschemaColumns[$TableName]
    }
    return @()
}

# --- Map TMSCHEMA DataType to SQL types ---
function Get-SqlType {
    param($DataType)
    switch ($DataType) {
        2    { return "NVARCHAR(MAX)" }   # String
        6    { return "BIGINT" }          # Int64 (Whole Number)
        8    { return "FLOAT" }           # Double (Decimal Number)
        9    { return "DATETIME" }        # DateTime
        10   { return "DECIMAL(18,4)" }   # Decimal (Currency)
        11   { return "BIT" }             # Boolean
        default { return "NVARCHAR(MAX)" }
    }
}

# --- Escape SQL string value ---
function Format-SqlValue {
    param($Value, $DataType)
    if ($null -eq $Value -or $Value.ToString() -eq '') {
        return "NULL"
    }
    $strVal = $Value.ToString()
    switch ($DataType) {
        { $_ -in 6, 8, 10 } {
            # Numeric (Int64, Double, Decimal)
            if ($strVal -match '^-?[\d.]+([eE][+-]?\d+)?$') { return $strVal }
            else { return "NULL" }
        }
        11 {
            # Boolean
            if ($strVal -eq 'True') { return "1" }
            else { return "0" }
        }
        9 {
            # DateTime
            try {
                $dt = [DateTime]::Parse($strVal)
                return "'" + $dt.ToString("yyyy-MM-dd HH:mm:ss") + "'"
            } catch {
                return "NULL"
            }
        }
        default {
            # String - escape single quotes
            return "N'" + $strVal.Replace("'", "''") + "'"
        }
    }
}

# --- Sanitize table name for SQL ---
function Get-SqlTableName {
    param([string]$Name)
    $clean = $Name -replace '\s*\(\d+\)\s*$', ''  # Remove (2) suffix
    $clean = $clean -replace '[&]', 'And'
    $clean = $clean -replace '[^A-Za-z0-9_]', '_'
    if ($clean -eq 'Table') { $clean = 'FactTable' }
    return $clean
}

# --- Ensure data folder for CSV export ---
$dataFolder = Join-Path $PSScriptRoot "data"
if (-not (Test-Path $dataFolder)) {
    New-Item -ItemType Directory -Path $dataFolder -Force | Out-Null
}

# --- Format a raw value for CSV output ---
function Format-CsvValue {
    param($Value, $DataType)
    if ($null -eq $Value -or $Value.ToString() -eq '') {
        return ''
    }
    $strVal = $Value.ToString()
    switch ($DataType) {
        9 {
            try {
                $dt = [DateTime]::Parse($strVal)
                return $dt.ToString("yyyy-MM-dd HH:mm:ss")
            } catch {
                return $strVal
            }
        }
        11 {
            if ($strVal -eq 'True') { return '1' } else { return '0' }
        }
        default {
            # Escape pipe and newline characters
            return $strVal.Replace('|', '\|').Replace("`r", '').Replace("`n", '\n')
        }
    }
}

# --- Build output ---
$sb = [System.Text.StringBuilder]::new()
[void]$sb.AppendLine("-- ============================================================")
[void]$sb.AppendLine("-- SQL INSERT statements generated from Power BI model data")
[void]$sb.AppendLine("-- Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$sb.AppendLine("-- ============================================================")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("SET DATEFORMAT ymd;")
[void]$sb.AppendLine("GO")
[void]$sb.AppendLine("")

foreach ($table in $tables) {
  try {
    Write-Host "`nProcessing table: $table"
    $sqlTableName = Get-SqlTableName $table
    $columns = Get-TableColumns $table
    if ($columns.Count -eq 0) {
        Write-Host "  Skipping (no columns)"
        continue
    }

    $colNames = $columns | ForEach-Object { "[$($_.Name)]" }
    $colNameList = $colNames -join ', '

    # --- Table header comment ---
    [void]$sb.AppendLine("-- ============================================================")
    [void]$sb.AppendLine("-- Table: $sqlTableName (source: $table)")
    [void]$sb.AppendLine("-- ============================================================")
    [void]$sb.AppendLine("")

    # --- Query data via DAX ---
    $dataCmd = $conn.CreateCommand()
    $escapedTable = $table.Replace("'", "''")
    $dataCmd.CommandText = "EVALUATE '$escapedTable'"
    try {
        $dataReader = $dataCmd.ExecuteReader()
    } catch {
        Write-Host "  Warning: Could not query table '$table': $_"
        continue
    }

    # --- Prepare CSV file ---
    $csvPath = Join-Path $dataFolder "$sqlTableName.csv"
    $csvSb = [System.Text.StringBuilder]::new()
    $csvHeader = ($columns | ForEach-Object { $_.Name }) -join '|'
    [void]$csvSb.AppendLine($csvHeader)

    # --- Build ordinal lookup once (avoid O(rows*cols*fields) inner loop) ---
    $ordinals = @()
    for ($i = 0; $i -lt $columns.Count; $i++) {
        $col = $columns[$i]
        $ord = -1
        for ($j = 0; $j -lt $dataReader.FieldCount; $j++) {
            if ($dataReader.GetName($j) -like "*$($col.Name)*") {
                $ord = $j
                break
            }
        }
        $ordinals += $ord
    }

    $rowCount = 0
    $batchRows = [System.Collections.Generic.List[string]]::new()

    while ($dataReader.Read()) {
        $values = [string[]]::new($columns.Count)
        $csvValues = [string[]]::new($columns.Count)
        for ($i = 0; $i -lt $columns.Count; $i++) {
            $ord = $ordinals[$i]
            if ($ord -ge 0 -and -not $dataReader.IsDBNull($ord)) {
                $val = $dataReader.GetValue($ord)
                $values[$i] = Format-SqlValue $val $columns[$i].DataType
                $csvValues[$i] = Format-CsvValue $val $columns[$i].DataType
            } else {
                $values[$i] = "NULL"
                $csvValues[$i] = ''
            }
        }
        [void]$batchRows.Add("($($values -join ', '))")
        [void]$csvSb.AppendLine($csvValues -join '|')
        $rowCount++

        if ($rowCount % 1000 -eq 0) {
            Write-Host "`r    ... $rowCount rows processed" -NoNewline
        }

        if ($batchRows.Count -ge $BatchSize) {
            [void]$sb.AppendLine("INSERT INTO [$sqlTableName] ($colNameList)")
            [void]$sb.AppendLine("VALUES")
            [void]$sb.AppendLine(($batchRows -join ",`n") + ";")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")
            $batchRows.Clear()
        }
    }

    # Flush remaining rows
    if ($batchRows.Count -gt 0) {
        [void]$sb.AppendLine("INSERT INTO [$sqlTableName] ($colNameList)")
        [void]$sb.AppendLine("VALUES")
        [void]$sb.AppendLine(($batchRows -join ",`n") + ";")
        [void]$sb.AppendLine("GO")
        [void]$sb.AppendLine("")
    }

    $dataReader.Close()

    # --- Write CSV file ---
    [System.IO.File]::WriteAllText($csvPath, $csvSb.ToString(), [System.Text.Encoding]::UTF8)
    Write-Host "`r    Exported $rowCount rows (CSV: $csvPath)          "
  } catch {
    Write-Host "  ERROR processing table '$table': $_"
  }
}

$conn.Close()

# --- Write output file ---
$outputPath = Join-Path $PSScriptRoot $OutputFile
[System.IO.File]::WriteAllText($outputPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
Write-Host "`n============================================================"
Write-Host "Done! SQL script written to: $outputPath"
Write-Host "CSV files written to: $dataFolder"
Write-Host "============================================================"
Write-Host ""
Write-Host "NOTE: Run 'python extract_schema.py' first to generate create_tables.sql"
