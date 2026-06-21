# DGE Report Comparator

This repository contains the DGE report comparator and its MCP server.

## Project Location

Use this folder for all future changes:

```text
C:\Users\YuliaFarber\Documents\dev\dge-report-comparator
```

## Comparator Source Of Truth

- Keep comparison behavior in `dge_compare.py`.
- Keep MCP wrapping/orchestration in `dge_mcp_server.py`.
- Do not duplicate comparison rules in a second implementation.
- If report matching, normalization, diff grouping, dashboard generation, or status rules need to change, update `dge_compare.py` first and let the MCP server delegate to it.

## MCP Usage

When the user says "report comparator", "DGE comparator", "compare reports", or asks to run the reconciliation, prefer the `dge_reconciliation` MCP tools when available:

- `run_dge_reconciliation`
- `run_caesars_reconciliation`
- `list_report_runs`
- `get_report_run`

Use `run_dge_reconciliation` for new calls. `run_caesars_reconciliation` exists for backward compatibility.

Common local compare examples:

```text
brand=caesars, source_mode=local, from_date=YYYYMMDD, until_date=YYYYMMDD
brand=fd, source_mode=local, from_date=YYYYMMDD, until_date=YYYYMMDD
brand=all, source_mode=local, from_date=YYYYMMDD, until_date=YYYYMMDD
```

The project-scoped MCP configuration is in `.codex/config.toml`. It starts:

```text
dge_mcp_server.py
```

using the repository-local virtual environment:

```text
.venv\Scripts\python.exe
```

## Setup

If the MCP server does not start, reinstall dependencies into the local virtual environment:

```powershell
& .\.venv\Scripts\python.exe -m pip install -r requirements-mcp.txt
```

If `.venv` does not exist:

```powershell
& C:\Users\YuliaFarber\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m venv .venv
& .\.venv\Scripts\python.exe -m pip install -r requirements-mcp.txt
```

## Verification

Before committing MCP changes, run:

```powershell
& .\.venv\Scripts\python.exe -c "import dge_compare; import dge_mcp_server; print('ok')"
```

For logic changes, run a local compare with known SFTP/Fabric folders and inspect `summary.json`, `details.json`, and `dashboard.html`.

## Safety

- Do not commit `.env`, `*.local.env`, downloaded report files, or `.venv`.
- Redact secrets from logs and examples.
- Treat PowerShell download scripts as orchestration; comparison semantics belong in `dge_compare.py`.
