# DGE Report Reconciliation

Compares the daily DGE XLSX reports produced from Mongo/SFTP with the matching
reports produced from Fabric.

The comparator checks workbook values, not raw XLSX bytes. This avoids false
failures caused by Excel metadata, workbook styles, compression differences, or
file creation timestamps.

## Reports

The supported report filename families are:

- `game_round_report_{brand}_{date}`
- `game_summary_report_{brand}_{date}`
- `machine_summary_report_{brand}_{date}`
- `pending_transaction_{brand}_{date}`
- `session_transaction_{brand}_{date}`
- `void-transaction_{brand}_{date}`
- `void_transaction_{brand}_{date}`

The parser is case-insensitive and accepts both `_` and `-` inside the report
type. The date is parsed as the last `yyyyMMdd` value in the filename.

Example:

```text
VOID_TRANSACTION_caesars-nj20260530.xlsx
```

is parsed as:

```text
report = void_transaction
brand = caesars-nj
date = 2026-05-30
```

## Quick Local Run

Download or copy the Mongo/SFTP files into one folder and the Fabric files into
another folder, then run:

```powershell
python tools\dge-reconciliation\dge_compare.py `
  --sftp-dir C:\tmp\dge\sftp `
  --fabric-dir C:\tmp\dge\fabric `
  --output-dir C:\tmp\dge\out `
  --brand caesars-nj `
  --report void_transaction `
  --date 20260530 `
  --date 20260529
```

Outputs:

- `dashboard.html` - local dashboard with date tabs, report filters, period
  status bars, report status bars, source-file links, mismatch details, and
  open-file links inside the detail view
- `summary.csv` / `summary.json` - one row per report/date/brand
- `details.csv` / `details.json` - mismatch details

By default, rows are compared as a multiset, so row order may differ but the
same rows and duplicate counts must exist on both sides. Add `--strict-order`
when row order must match exactly.

Omit `--report` to check all six DGE report families.

When a row is not exactly equal, the comparator first tries to pair it with the
most similar row on the other side. This turns simple edits into field-level
differences, for example `Rounds: 4236 -> 4237`, instead of showing one
`missing_from_fabric` row and one `extra_in_fabric` row.

In the dashboard, row/value differences are grouped by matched row. If the same
row has several changed cells, it is shown once with a note such as
`4 changed fields in this same row`, followed by the individual changed fields.
Rows are grouped by an internal matched-row id, so two different rows with weak
identifiers, for example only the same date and operator, are not merged into
one dashboard group.

## Caesars Pilot

Fabric source folder:

```text
https://app.fabric.microsoft.com/groups/912d99dd-0947-4c54-ab3b-166a11cf0f0e/lakehouses/f4d2b03c-4a71-42ac-887b-6394865db591?experience=power-bi&selectedPath=Files%2Fauto_reports%2Fregulatory%2Fcaesars-nj&extensionScenario=openArtifact
```

Azure SFTP / storage portal source:

```text
https://portal.azure.com/#view/Microsoft_Azure_Storage/ContainerMenuBlade/~/overview/storageAccountId/%2Fsubscriptions%2Ffcb6ca50-8dc2-4c83-b963-1734f7f053ec%2FresourceGroups%2Fawager%2Fproviders%2FMicrosoft.Storage%2FstorageAccounts%2Fawagersftp/path/c8-nj-prod/etag/%220x8DD11EBF1A834F9%22/defaultId//publicAccessVal/None
```

Current rule for the pilot:

- Brands:
  - Caesars: compare brand `caesars-nj`; SFTP brand `caesars-nj`; Fabric
    brand/folder `caesars-nj`.
  - FD / FanDuel: compare brand `fd`; SFTP brand/folder `fd`; Fabric
    brand/folder `fanduel-nj`.
- Reports checked: `game_summary_report`, `machine_summary_report`,
  `pending_transaction`, `void_transaction`.
- Reports excluded for now: `game_round_report`, `session_transaction`.
- Filename date means the previous day's data.
- UTC vs EDT timestamps are considered a mismatch for now because the two
  exports should be the same.
- The dashboard includes links to both source folders and to both local file
  copies for every result row, whether the result is green or red.

The scripts accept `-Brand caesars`, `-Brand caesars-nj`, `-Brand fd`, or
`-Brand fanduel-nj`. For FD, the comparison automatically maps Fabric
`fanduel-nj` files to logical brand `fd`.

## Local Azure Blob Download

Use this path when DevOps grants the App Registration access to the storage
container. It reads the Azure Storage blobs directly instead of using SFTP.

Known App Registration:

```text
Application name: dge-report-reader
Client ID: 28aa5709-ef31-48d7-b919-c858650b52af
Storage account: awagersftp
Container: c8-nj-prod
```

Required Azure role:

```text
Preferred: Storage Blob Data Reader
Also works but broader: Storage Blob Data Contributor
```

Ask DevOps for:

```text
AZURE_TENANT_ID
AZURE_CLIENT_SECRET or certificate details
DGE_STORAGE_PREFIX exact folder/prefix for Caesars files
```

Install Azure CLI locally if needed, then create:

```text
C:\tmp\dge\caesars\.env
```

With:

```text
AZURE_TENANT_ID=<tenant-id>
AZURE_CLIENT_ID=28aa5709-ef31-48d7-b919-c858650b52af
AZURE_CLIENT_SECRET=<secret>

DGE_STORAGE_ACCOUNT=awagersftp
DGE_STORAGE_CONTAINER=c8-nj-prod
DGE_STORAGE_PREFIX=<caesars folder/prefix>
DGE_BLOB_LOCAL_DIR=C:\tmp\dge\caesars\sftp

FABRIC_WORKSPACE_ID=912d99dd-0947-4c54-ab3b-166a11cf0f0e
FABRIC_LAKEHOUSE_ID=f4d2b03c-4a71-42ac-887b-6394865db591
FABRIC_FOLDER=Files/auto_reports/regulatory/caesars-nj
DGE_FABRIC_DIR=C:\tmp\dge\caesars\fabric
DGE_OUTPUT_DIR=C:\tmp\dge\caesars\out
```

If Fabric uses the same service principal as Azure Blob/SFTP, leave
`FABRIC_TENANT_ID`, `FABRIC_CLIENT_ID`, and `FABRIC_CLIENT_SECRET` empty. The
script reuses `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, and `AZURE_CLIENT_SECRET`.

Then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\dge-reconciliation\run_caesars_blob_local.ps1 `
  -EnvFile C:\tmp\dge\caesars\.env `
  -FromDate 20260528 `
  -UntilDate 20260530 `
  -OpenDashboard
```

`-FromDate` and `-UntilDate` are inclusive filename/report dates. The script
first downloads the four SFTP/Mongo report files for every date in the range
into `DGE_BLOB_LOCAL_DIR`, then downloads the four Fabric report files into
`DGE_FABRIC_DIR`, then compares only the same date range.

By default, files that already exist in `DGE_BLOB_LOCAL_DIR` and
`DGE_FABRIC_DIR` are reused and not downloaded again. To re-download and
overwrite existing local files on both sides, add `-ForceDownload`:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\dge-reconciliation\run_caesars_blob_local.ps1 `
  -EnvFile C:\tmp\dge\caesars\.env `
  -FromDate 20260528 `
  -UntilDate 20260530 `
  -ForceDownload `
  -OpenDashboard
```

The Caesars pilot comparison is limited to these four report families:

```text
game_summary_report
machine_summary_report
pending_transaction
void_transaction
```

If you prefer "last N days" instead of an explicit range, use `-EndDate` and
`-Days`:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\dge-reconciliation\run_caesars_blob_local.ps1 `
  -EnvFile C:\tmp\dge\caesars\.env `
  -EndDate 20260530 `
  -Days 3 `
  -OpenDashboard
```

## Local Fabric OneLake Download

To test only the Fabric download, use:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\dge-reconciliation\download_caesars_fabric.ps1 `
  -EnvFile C:\tmp\dge\caesars\.env `
  -WorkspaceId 912d99dd-0947-4c54-ab3b-166a11cf0f0e `
  -LakehouseId f4d2b03c-4a71-42ac-887b-6394865db591 `
  -Folder Files/auto_reports/regulatory/caesars-nj `
  -LocalDir C:\tmp\dge\caesars\fabric `
  -FromDate 20260528 `
  -UntilDate 20260530
```

Add `-ForceDownload` to re-download and overwrite existing Fabric files.

## FD / FanDuel

For FD / FanDuel, use the same scripts with `-Brand fd`. The defaults are:

```text
SFTP/blob brand: fd
SFTP/blob prefix: fd
SFTP remote folder: /c8-nj-prod/fd
Fabric brand: fanduel-nj
Fabric folder: Files/auto_reports/regulatory/fanduel-nj
Local SFTP folder: C:\tmp\dge\fd\sftp
Local Fabric folder: C:\tmp\dge\fd\fabric
Output folder: C:\tmp\dge\fd\out
```

Local FD compare:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\dge-reconciliation\run_caesars_local.ps1 `
  -Brand fd `
  -SftpDir C:\tmp\dge\fd\sftp `
  -FabricDir C:\tmp\dge\fd\fabric `
  -OutputDir C:\tmp\dge\fd\out `
  -FromDate 20260501 `
  -UntilDate 20260531 `
  -OpenDashboard
```

FD download from Azure Blob/Fabric and then compare:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\dge-reconciliation\run_caesars_blob_local.ps1 `
  -EnvFile C:\tmp\dge\caesars\.env `
  -Brand fd `
  -FromDate 20260501 `
  -UntilDate 20260531 `
  -OpenDashboard
```

If the real SFTP FD folder is not `/c8-nj-prod/fd`, override it:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\dge-reconciliation\run_caesars_sftp_local.ps1 `
  -EnvFile C:\tmp\dge\caesars\.env `
  -Brand fd `
  -SftpRemoteDir /c8-nj-prod/<actual-fd-folder> `
  -FromDate 20260501 `
  -UntilDate 20260531
```

## Local SFTP Download

Use this path when Fabric files are local, but Mongo/SFTP files should be pulled
from Azure SFTP automatically.

Azure Blob SFTP with local users uses SSH key authentication for automation.
This is not the same as Azure RBAC service-principal access.

### 1. Create an SSH key

```powershell
mkdir $env:USERPROFILE\.ssh -Force

ssh-keygen -t rsa -b 4096 `
  -f $env:USERPROFILE\.ssh\dge_caesars_sftp `
  -C "dge-caesars-sftp-reader"
```

This creates:

```text
C:\Users\YuliaFarber\.ssh\dge_caesars_sftp
C:\Users\YuliaFarber\.ssh\dge_caesars_sftp.pub
```

Keep the private key secret. Give only the `.pub` public key to the Azure admin.

### 2. Ask Azure admin to create the SFTP local user

Request:

```text
Storage account: awagersftp
SFTP local user: dge-caesars-reader
Authentication: SSH public key
Public key: content of C:\Users\YuliaFarber\.ssh\dge_caesars_sftp.pub
Permission: read/list only
Scope: c8-nj-prod and the Caesars report path if folder scoping is available
```

The Azure SFTP username usually has this shape:

```text
awagersftp.dge-caesars-reader
```

### 3. Test SFTP manually

```powershell
sftp -i $env:USERPROFILE\.ssh\dge_caesars_sftp awagersftp.dge-caesars-reader@awagersftp.blob.core.windows.net
```

Inside SFTP:

```text
ls
cd c8-nj-prod
ls
bye
```

Adjust the remote folder until you see the Caesars XLSX files.

### 4. Create a local env file

Copy `.env.example` to:

```text
C:\tmp\dge\caesars\.env
```

Fill:

```text
DGE_SFTP_HOST=awagersftp.blob.core.windows.net
DGE_SFTP_PORT=22
DGE_SFTP_USERNAME=awagersftp.dge-caesars-reader
DGE_SFTP_PRIVATE_KEY=C:\Users\YuliaFarber\.ssh\dge_caesars_sftp
DGE_SFTP_REMOTE_DIR=/c8-nj-prod

DGE_SFTP_LOCAL_DIR=C:\tmp\dge\caesars\sftp
DGE_FABRIC_DIR=C:\tmp\dge\caesars\fabric
DGE_OUTPUT_DIR=C:\tmp\dge\caesars\out
```

### 5. Download from SFTP and compare

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\dge-reconciliation\run_caesars_sftp_local.ps1 `
  -EnvFile C:\tmp\dge\caesars\.env `
  -FromDate 20260528 `
  -UntilDate 20260530 `
  -OpenDashboard
```

`-FromDate` and `-UntilDate` are inclusive filename/report dates. The script
first downloads the four SFTP/Mongo report files for every date in the range
into `DGE_SFTP_LOCAL_DIR`. If Fabric OneLake settings are configured, it also
downloads the four Fabric report files into `DGE_FABRIC_DIR`, then compares only
the same date range.

By default, files that already exist in `DGE_SFTP_LOCAL_DIR` and
`DGE_FABRIC_DIR` are reused and not downloaded again. To re-download and
overwrite existing local files on both sides, add `-ForceDownload`.

If you prefer "last N days" instead of an explicit range, use `-EndDate` and
`-Days`:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\dge-reconciliation\run_caesars_sftp_local.ps1 `
  -EnvFile C:\tmp\dge\caesars\.env `
  -EndDate 20260530 `
  -Days 3 `
  -OpenDashboard
```

After downloading the files locally, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\dge-reconciliation\run_caesars_local.ps1 `
  -SftpDir C:\tmp\dge\caesars\sftp `
  -FabricDir C:\tmp\dge\caesars\fabric `
  -OutputDir C:\tmp\dge\caesars\out `
  -FromDate 20260528 `
  -UntilDate 20260530 `
  -OpenDashboard
```

The PowerShell command exits with code `0` when all checks are green and a
non-zero code when at least one report is red or missing. The dashboard is still
written in both cases.

If you prefer "last N days" instead of an explicit range:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\dge-reconciliation\run_caesars_local.ps1 `
  -SftpDir C:\tmp\dge\caesars\sftp `
  -FabricDir C:\tmp\dge\caesars\fabric `
  -OutputDir C:\tmp\dge\caesars\out `
  -EndDate 20260530 `
  -Days 3 `
  -OpenDashboard
```

## MCP Server

The same reconciliation can be exposed as an MCP server, so Codex or another
MCP client can run comparisons and read previous dashboard runs without calling
the PowerShell scripts manually.

Install the MCP dependencies in the Python environment that will run the server:

```powershell
cd C:\Users\YuliaFarber\Documents\dev\report\tools\dge-reconciliation
& C:\Users\YuliaFarber\AppData\Local\Python\pythoncore-3.14-64\python.exe -m pip install -r requirements-mcp.txt
```

Example MCP client configuration:

```json
{
  "mcpServers": {
    "dge-reconciliation": {
      "command": "C:\\Users\\YuliaFarber\\AppData\\Local\\Python\\pythoncore-3.14-64\\python.exe",
      "args": [
        "C:\\Users\\YuliaFarber\\Documents\\dev\\report\\tools\\dge-reconciliation\\dge_mcp_server.py"
      ]
    }
  }
}
```

The server exposes these tools:

- `run_dge_reconciliation` - runs DGE reconciliation for `brand=caesars`,
  `brand=fd`, or `brand=all`.
- `run_caesars_reconciliation` - runs the Caesars reconciliation and returns
  green/red status counts, dashboard links, summary rows, and a diff preview.
  This is kept for backward compatibility; prefer `run_dge_reconciliation`.
  Use `source_mode=local` for already downloaded folders,
  `source_mode=blob` for Azure Blob API plus Fabric download, or
  `source_mode=sftp` for Azure Storage SFTP plus Fabric download.
- `list_report_runs` - lists archived dashboards in the output folder.
- `get_report_run` - reads one archived run and returns its summary and diffs.

Example MCP tool arguments for a local compare:

```json
{
  "brand": "caesars",
  "source_mode": "local",
  "sftp_dir": "C:\\tmp\\dge\\caesars\\sftp",
  "fabric_dir": "C:\\tmp\\dge\\caesars\\fabric",
  "output_dir": "C:\\tmp\\dge\\caesars\\out",
  "from_date": "20260528",
  "until_date": "20260530"
}
```

Example MCP tool arguments for FD:

```json
{
  "brand": "fd",
  "source_mode": "local",
  "from_date": "20260501",
  "until_date": "20260531"
}
```

Example MCP tool arguments for both supported brands:

```json
{
  "brand": "all",
  "source_mode": "local",
  "from_date": "20260501",
  "until_date": "20260531"
}
```

Example MCP tool arguments for service-principal download plus compare:

```json
{
  "source_mode": "blob",
  "env_file": "C:\\tmp\\dge\\caesars\\.env",
  "from_date": "20260528",
  "until_date": "20260530",
  "force_download": false
}
```

## Created-Time Filter

If the input folders contain multiple copies, filter by local file modified time:

```powershell
python tools\dge-reconciliation\dge_compare.py `
  --sftp-dir C:\tmp\dge\sftp `
  --fabric-dir C:\tmp\dge\fabric `
  --output-dir C:\tmp\dge\out `
  --brand caesars-nj `
  --days 2 `
  --created-from 12:00 `
  --created-to 13:00
```

For production Azure/Fabric connectors, prefer the remote object `LastModified`
timestamp instead of the local downloaded file timestamp.

## Production Wiring

To automate this daily, the service needs programmatic access to both locations:

1. Azure SFTP storage account/container/path for the Mongo-produced files.
2. Fabric Lakehouse OneLake/workspace/lakehouse path for `Files/auto_reports/regulatory/caesars-nj`.
3. One service principal or managed identity with read access to both locations.
4. A scheduled run after both report-generation jobs complete, for example
   daily after 13:00 in the chosen operations timezone.
5. A durable result table or folder for `summary` and `details` outputs.
6. Alerts on `MISSING_*` or `MISMATCH` statuses.

Recommended dashboard source fields:

- business date
- brand
- report
- status
- SFTP/Fabric paths
- SFTP/Fabric row counts
- SFTP/Fabric column counts
- difference count
- checked timestamp

Power BI can read the emitted CSV/JSON outputs, or the same fields can be
persisted into Fabric/SQL for a longer historical dashboard.
