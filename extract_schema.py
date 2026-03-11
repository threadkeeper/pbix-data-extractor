import re
import json
import os
import subprocess
import sys

# =============================================================
# extract_schema.py
# Connects to the local Power BI Desktop SSAS model to get
# the full table/column schema, then generates create_tables.sql.
# Falls back to parsing the .pbix layout if SSAS is not available.
# =============================================================

script_dir = os.path.dirname(os.path.abspath(__file__))

def sanitize_table_name(name):
    clean = re.sub(r'\s*\(\d+\)\s*$', '', name)  # Remove (2) suffix
    clean = clean.replace('&', 'And')
    clean = re.sub(r'[^A-Za-z0-9_]', '_', clean)
    if clean == 'Table':
        clean = 'FactTable'
    return clean

# --- SSAS data type to SQL type mapping ---
SSAS_TYPE_MAP = {
    5:   'DECIMAL(18,2)',   # DBTYPE_R8 / Double
    6:   'DECIMAL(18,4)',   # DBTYPE_CY / Currency
    7:   'DATETIME',        # DBTYPE_DATE
    11:  'BIT',             # DBTYPE_BOOL
    20:  'BIGINT',          # DBTYPE_I8
    3:   'INT',             # DBTYPE_I4
    130: 'NVARCHAR(MAX)',   # DBTYPE_WSTR
}

def get_sql_type_from_ssas(dtype):
    return SSAS_TYPE_MAP.get(dtype, 'NVARCHAR(MAX)')

# --- Try connecting to SSAS via PowerShell ---
PS_SCRIPT = r'''
Add-Type -Path "C:\Program Files\Microsoft.NET\ADOMD.NET\170\Microsoft.AnalysisServices.AdomdClient.dll"
$msmdsrv = Get-Process -Name "msmdsrv" -ErrorAction SilentlyContinue
if (-not $msmdsrv) { Write-Error "NOT_RUNNING"; exit 1 }
$port = (Get-NetTCPConnection -OwningProcess $msmdsrv.Id -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1).LocalPort
$conn = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection("Data Source=localhost:$port")
$conn.Open()

# First get the list of user tables
$cmd = $conn.CreateCommand()
$cmd.CommandText = "SELECT [TABLE_NAME] FROM `$SYSTEM.DBSCHEMA_TABLES WHERE TABLE_TYPE = 'TABLE'"
$reader = $cmd.ExecuteReader()
$userTables = @()
while ($reader.Read()) { $userTables += $reader[0].ToString() }
$reader.Close()

# Now get columns only for user tables
$cmd2 = $conn.CreateCommand()
$cmd2.CommandText = "SELECT [TABLE_NAME], [COLUMN_NAME], [DATA_TYPE] FROM `$SYSTEM.DBSCHEMA_COLUMNS"
$reader2 = $cmd2.ExecuteReader()
$results = @()
while ($reader2.Read()) {
    $tbl = $reader2[0].ToString()
    if ($userTables -contains $tbl) {
        $col = $reader2[1].ToString()
        $dt = $reader2[2]
        $results += "$tbl`t$col`t$dt"
    }
}
$reader2.Close()
$conn.Close()
$results -join "`n"
'''

def query_ssas_schema():
    """Query the live SSAS model for full table/column schema."""
    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', PS_SCRIPT],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode != 0:
            return None
        tables = {}
        for line in result.stdout.strip().split('\n'):
            parts = line.strip().split('\t')
            if len(parts) != 3:
                continue
            raw_table, col_name, dtype_str = parts
            # DMV returns table names with $ prefix
            table = raw_table.lstrip('$')
            # Skip internal tables and RowNumber columns
            if table.startswith('LocalDateTable') or table.startswith('DateTableTemplate'):
                continue
            if col_name.startswith('RowNumber'):
                continue
            try:
                dtype = int(dtype_str)
            except ValueError:
                dtype = 130
            if table not in tables:
                tables[table] = []
            tables[table].append({'name': col_name, 'sql_type': get_sql_type_from_ssas(dtype)})
        return tables
    except Exception as e:
        print(f"Could not connect to SSAS: {e}")
        return None

def infer_sql_type(col_name):
    """Fallback type inference from column name."""
    lower = col_name.lower()
    if any(k in lower for k in ['date', 'createdat', 'updatedat']):
        return 'DATETIME'
    if any(k in lower for k in ['amount', 'total', 'payment', 'partpayment']):
        return 'DECIMAL(18,2)'
    if any(k in lower for k in ['age', 'days', 'order']):
        return 'INT'
    if lower == 'id' or lower.endswith('id'):
        return 'INT'
    return 'NVARCHAR(MAX)'

# --- Try SSAS first ---
print("Connecting to Power BI Desktop SSAS model...")
ssas_schema = query_ssas_schema()

if ssas_schema:
    print(f"Connected! Found {len(ssas_schema)} tables from the live model.\n")
    source = 'Power BI SSAS model'
    table_data = {}
    for table, cols in ssas_schema.items():
        sql_name = sanitize_table_name(table)
        table_data[sql_name] = {
            'source': table,
            'columns': [(c['name'], c['sql_type']) for c in cols]
        }
else:
    print("SSAS not available. Falling back to .pbix layout parsing...\n")
    source = 'Power BI report layout (visual-referenced columns only)'
    layout_path = os.path.join(script_dir, 'pbix_extracted', 'Report', 'Layout')
    if not os.path.exists(layout_path):
        print("ERROR: pbix_extracted folder not found. Extract the .pbix first:")
        print('  Copy-Item "YourReport.pbix" "pbix_extracted.zip"')
        print('  Expand-Archive "pbix_extracted.zip" -DestinationPath "pbix_extracted" -Force')
        print('  Remove-Item "pbix_extracted.zip"')
        sys.exit(1)
    with open(layout_path, 'rb') as f:
        content = f.read().decode('utf-16-le')
    unescaped = content
    for _ in range(5):
        unescaped = unescaped.replace('\\"', '"').replace('\\\\', '\\')
    table_columns = {}
    for match in re.finditer(r'"Entity"\s*:\s*"([^"]+)"(.{0,200})"Property"\s*:\s*"([^"]+)"', unescaped):
        entity, prop = match.group(1), match.group(3)
        table_columns.setdefault(entity, set()).add(prop)
    for match in re.findall(r"'([^']+)'\[([^\]]+)\]", unescaped):
        table_columns.setdefault(match[0], set()).add(match[1])
    skip = ('LocalDateTable', 'DateTableTemplate', 'MCSB')
    table_data = {}
    for table in sorted(table_columns.keys()):
        if any(table.startswith(p) for p in skip):
            continue
        cols = sorted(table_columns[table])
        if not cols:
            continue
        sql_name = sanitize_table_name(table)
        table_data[sql_name] = {
            'source': table,
            'columns': [(c, infer_sql_type(c)) for c in cols]
        }

# --- Print summary ---
print("=" * 60)
print(f"SOURCE: {source}")
print("=" * 60)
for sql_name in sorted(table_data.keys()):
    info = table_data[sql_name]
    print(f"\nTable: {sql_name} (source: {info['source']})")
    for col_name, col_type in info['columns']:
        print(f"    {col_name:40s} {col_type}")
print(f"\nTotal tables: {len(table_data)}")
print(f"Total columns: {sum(len(v['columns']) for v in table_data.values())}")

# --- Generate create_tables.sql ---
output_path = os.path.join(script_dir, 'create_tables.sql')
with open(output_path, 'w', encoding='utf-8') as f:
    f.write('-- ============================================================\n')
    f.write(f'-- SQL CREATE TABLE statements\n')
    f.write(f'-- Source: {source}\n')
    f.write('-- ============================================================\n\n')
    for sql_name in sorted(table_data.keys()):
        info = table_data[sql_name]
        f.write(f"IF OBJECT_ID('{sql_name}', 'U') IS NOT NULL DROP TABLE [{sql_name}];\n")
        f.write('GO\n')
        f.write(f'CREATE TABLE [{sql_name}] (\n')
        cols = info['columns']
        for i, (col_name, col_type) in enumerate(cols):
            comma = ',' if i < len(cols) - 1 else ''
            f.write(f'    [{col_name}] {col_type}{comma}\n')
        f.write(');\n')
        f.write('GO\n\n')

print(f"\nGenerated: {output_path}")
