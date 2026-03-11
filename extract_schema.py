import re
import os
import subprocess
import sys

# =============================================================
# extract_schema.py
# Connects to the local Power BI Desktop SSAS model to get
# the full table/column schema, then generates create_tables.sql.
# Uses TMSCHEMA_COLUMNS for accurate data types.
# Requires Power BI Desktop to be open with the report loaded.
# =============================================================

script_dir = os.path.dirname(os.path.abspath(__file__))

def sanitize_table_name(name):
    clean = re.sub(r'\s*\(\d+\)\s*$', '', name)  # Remove (2) suffix
    clean = clean.replace('&', 'And')
    clean = re.sub(r'[^A-Za-z0-9_]', '_', clean)
    if clean == 'Table':
        clean = 'FactTable'
    return clean

# --- TMSCHEMA DataType to SQL type mapping ---
# TMSCHEMA_COLUMNS.DataType values:
#   2  = String
#   6  = Int64 (Whole Number)
#   8  = Double (Decimal Number)
#   9  = DateTime
#   10 = Decimal (Fixed Decimal / Currency)
#   11 = Boolean
TMSCHEMA_TYPE_MAP = {
    1:  'NVARCHAR(MAX)',    # Variant / calculated table string
    2:  'NVARCHAR(MAX)',
    6:  'BIGINT',
    8:  'FLOAT',
    9:  'DATETIME',
    10: 'DECIMAL(18,4)',
    11: 'BIT',
}

def get_sql_type_from_tmschema(dtype):
    return TMSCHEMA_TYPE_MAP.get(dtype, 'NVARCHAR(MAX)')

# --- Connect to SSAS via PowerShell ---
PS_SCRIPT = r'''
Add-Type -Path "C:\Program Files\Microsoft.NET\ADOMD.NET\170\Microsoft.AnalysisServices.AdomdClient.dll"
$msmdsrv = Get-Process -Name "msmdsrv" -ErrorAction SilentlyContinue
if (-not $msmdsrv) { Write-Error "NOT_RUNNING"; exit 1 }
$port = (Get-NetTCPConnection -OwningProcess $msmdsrv.Id -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1).LocalPort
$conn = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection("Data Source=localhost:$port")
$conn.Open()

# Get user tables from DBSCHEMA
$cmd = $conn.CreateCommand()
$cmd.CommandText = "SELECT [TABLE_NAME] FROM `$SYSTEM.DBSCHEMA_TABLES WHERE TABLE_TYPE = 'TABLE'"
$reader = $cmd.ExecuteReader()
$userTables = @()
while ($reader.Read()) { $userTables += $reader[0].ToString().TrimStart('$') }
$reader.Close()

# Get table ID -> Name mapping from TMSCHEMA_TABLES
$cmdTbl = $conn.CreateCommand()
$cmdTbl.CommandText = "SELECT [ID], [Name] FROM `$SYSTEM.TMSCHEMA_TABLES"
$tblReader = $cmdTbl.ExecuteReader()
$tableIdToName = @{}
while ($tblReader.Read()) {
    $tableIdToName[$tblReader[0].ToString()] = $tblReader[1].ToString()
}
$tblReader.Close()

# Get all data columns with real ExplicitDataType from TMSCHEMA_COLUMNS
# Type: 1 = Data column, 2 = Calculated, 3 = RowNumber, 4 = CalculatedTableColumn
# Only Type=1 ensures calculated tables (all Type=4 cols) are excluded entirely
$cmdCols = $conn.CreateCommand()
$cmdCols.CommandText = "SELECT [TableID], [ExplicitName], [ExplicitDataType] FROM `$SYSTEM.TMSCHEMA_COLUMNS WHERE [Type] = 1"
$colReader = $cmdCols.ExecuteReader()
$results = @()
while ($colReader.Read()) {
    $tableId = $colReader[0].ToString()
    $colName = $colReader[1].ToString()
    $dataType = $colReader[2]
    if ($colName -eq '') { continue }
    $tableName = $tableIdToName[$tableId]
    if ($userTables -contains $tableName) {
        $results += "$tableName`t$colName`t$dataType"
    }
}
$colReader.Close()
$conn.Close()
$results -join "`n"
'''

print("Connecting to Power BI Desktop SSAS model...")
result = subprocess.run(
    ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', PS_SCRIPT],
    capture_output=True, text=True, timeout=30
)
if result.returncode != 0:
    print("ERROR: Could not connect to SSAS.")
    print("Make sure Power BI Desktop is open with the report loaded.")
    if result.stderr.strip():
        print(f"Detail: {result.stderr.strip()}")
    sys.exit(1)

tables = {}
for line in result.stdout.strip().split('\n'):
    parts = line.strip().split('\t')
    if len(parts) != 3:
        continue
    table, col_name, dtype_str = parts
    if table.startswith('LocalDateTable') or table.startswith('DateTableTemplate'):
        continue
    try:
        dtype = int(dtype_str)
    except ValueError:
        dtype = 2
    if table not in tables:
        tables[table] = []
    tables[table].append((col_name, get_sql_type_from_tmschema(dtype)))

print(f"Connected! Found {len(tables)} tables.\n")

# --- Print summary ---
source = 'Power BI SSAS model (TMSCHEMA)'
print("=" * 60)
print(f"SOURCE: {source}")
print("=" * 60)
for table in sorted(tables.keys()):
    sql_name = sanitize_table_name(table)
    print(f"\nTable: {sql_name} (source: {table})")
    for col_name, col_type in tables[table]:
        print(f"    {col_name:40s} {col_type}")
print(f"\nTotal tables: {len(tables)}")
print(f"Total columns: {sum(len(v) for v in tables.values())}")

# --- Generate create_tables.sql ---
output_path = os.path.join(script_dir, 'create_tables.sql')
with open(output_path, 'w', encoding='utf-8') as f:
    f.write('-- ============================================================\n')
    f.write(f'-- SQL CREATE TABLE statements\n')
    f.write(f'-- Source: {source}\n')
    f.write('-- ============================================================\n\n')
    for table in sorted(tables.keys()):
        sql_name = sanitize_table_name(table)
        f.write(f"IF OBJECT_ID('{sql_name}', 'U') IS NOT NULL DROP TABLE [{sql_name}];\n")
        f.write('GO\n')
        f.write(f'CREATE TABLE [{sql_name}] (\n')
        cols = tables[table]
        for i, (col_name, col_type) in enumerate(cols):
            comma = ',' if i < len(cols) - 1 else ''
            f.write(f'    [{col_name}] {col_type}{comma}\n')
        f.write(');\n')
        f.write('GO\n\n')

print(f"\nGenerated: {output_path}")
