#!/usr/bin/env python3
"""MCP server for running DGE report reconciliation.

The server intentionally delegates comparison and download behavior to the
existing PowerShell/Python scripts, so there is one source of truth for report
matching, dashboard generation, and output formats.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

try:
    from mcp.server.fastmcp import FastMCP
except ImportError as exc:  # pragma: no cover - depends on local MCP install.
    raise SystemExit(
        "The MCP Python SDK is not installed. Run: "
        "python -m pip install -r requirements-mcp.txt"
    ) from exc


ROOT = Path(__file__).resolve().parent
ONELAKE_WORKSPACE_ID = "912d99dd-0947-4c54-ab3b-166a11cf0f0e"
ONELAKE_LAKEHOUSE_ID = "f4d2b03c-4a71-42ac-887b-6394865db591"
STORAGE_SOURCE_URL = (
    "https://portal.azure.com/#view/Microsoft_Azure_Storage/ContainerMenuBlade/~/overview/"
    "storageAccountId/%2Fsubscriptions%2Ffcb6ca50-8dc2-4c83-b963-1734f7f053ec%2FresourceGroups%2Fawager"
    "%2Fproviders%2FMicrosoft.Storage%2FstorageAccounts%2Fawagersftp/path/c8-nj-prod/etag/%220x8DD11EBF1A834F9%22/defaultId//publicAccessVal/None"
)
DGE_REPORTS = (
    "game_summary_report",
    "machine_summary_report",
    "pending_transaction",
    "void_transaction",
)
DATE_RE = re.compile(r"^\d{8}$")
RUN_DASHBOARD_RE = re.compile(r"Wrote run dashboard:\s*(.+)")
LATEST_DASHBOARD_RE = re.compile(r"Wrote latest dashboard:\s*(.+)")
SECRET_RE = re.compile(r"(?i)(client_secret|secret|password|private_key)(\s*[=:]\s*)([^\s]+)")

mcp = FastMCP("dge-reconciliation")


@dataclass(frozen=True)
class BrandConfig:
    key: str
    aliases: tuple[str, ...]
    compare_brand: str
    sftp_brand: str
    fabric_brand: str
    slug: str
    default_sftp_dir: Path
    default_fabric_dir: Path
    default_output_dir: Path
    default_fabric_folder: str
    default_storage_prefix: str
    default_sftp_remote_dir: str


def _fabric_source_url(folder: str) -> str:
    selected_path = folder.replace("/", "%2F")
    return (
        f"https://app.fabric.microsoft.com/groups/{ONELAKE_WORKSPACE_ID}/lakehouses/{ONELAKE_LAKEHOUSE_ID}"
        f"?experience=power-bi&selectedPath={selected_path}&extensionScenario=openArtifact"
    )


BRAND_CONFIGS = {
    "caesars": BrandConfig(
        key="caesars",
        aliases=("caesars", "caesars-nj"),
        compare_brand="caesars-nj",
        sftp_brand="caesars-nj",
        fabric_brand="caesars-nj",
        slug="caesars",
        default_sftp_dir=Path(r"C:\tmp\dge\caesars\sftp"),
        default_fabric_dir=Path(r"C:\tmp\dge\caesars\fabric"),
        default_output_dir=Path(r"C:\tmp\dge\caesars\out"),
        default_fabric_folder="Files/auto_reports/regulatory/caesars-nj",
        default_storage_prefix="",
        default_sftp_remote_dir="/c8-nj-prod",
    ),
    "fd": BrandConfig(
        key="fd",
        aliases=("fd", "fanduel", "fanduel-nj"),
        compare_brand="fd",
        sftp_brand="fd",
        fabric_brand="fanduel-nj",
        slug="fd",
        default_sftp_dir=Path(r"C:\tmp\dge\fd\sftp"),
        default_fabric_dir=Path(r"C:\tmp\dge\fd\fabric"),
        default_output_dir=Path(r"C:\tmp\dge\fd\out"),
        default_fabric_folder="Files/auto_reports/regulatory/fanduel-nj",
        default_storage_prefix="fd",
        default_sftp_remote_dir="/c8-nj-prod/fd",
    ),
}
BRAND_ALIASES = {
    alias: config.key
    for config in BRAND_CONFIGS.values()
    for alias in config.aliases
}
DEFAULT_OUTPUT_DIR = BRAND_CONFIGS["caesars"].default_output_dir


def _read_dotenv(env_file: str | None) -> dict[str, str]:
    if not env_file:
        return {}

    path = Path(env_file).expanduser()
    if not path.exists():
        raise ValueError(f"Env file not found: {path}")

    values: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        name, value = line.split("=", 1)
        values[name.strip()] = value.strip().strip('"')
    return values


def _first_value(*values: str | os.PathLike[str] | None) -> str | None:
    for value in values:
        if value:
            return str(value)
    return None


def _resolve_path(*values: str | os.PathLike[str] | None) -> Path:
    value = _first_value(*values)
    if not value:
        raise ValueError("A required path value is missing.")
    return Path(os.path.expandvars(value)).expanduser()


def _select_brand_configs(brand: str) -> list[BrandConfig]:
    requested = [part.strip().lower() for part in brand.split(",") if part.strip()]
    if not requested:
        requested = ["caesars"]
    if any(part == "all" for part in requested):
        return list(BRAND_CONFIGS.values())

    configs: list[BrandConfig] = []
    seen: set[str] = set()
    for item in requested:
        key = BRAND_ALIASES.get(item, item)
        config = BRAND_CONFIGS.get(key)
        if not config:
            supported = ", ".join(sorted({*BRAND_CONFIGS.keys(), "all"}))
            raise ValueError(f"Unsupported brand '{item}'. Supported values: {supported}.")
        if config.key not in seen:
            configs.append(config)
            seen.add(config.key)
    return configs


def _brand_env(env_values: dict[str, str], config: BrandConfig, name: str) -> str | None:
    brand_key = f"DGE_{config.key.upper()}_{name}"
    return env_values.get(brand_key) or os.environ.get(brand_key)


def _generic_env(env_values: dict[str, str], name: str) -> str | None:
    return env_values.get(name) or os.environ.get(name)


def _file_uri(path: Path | None) -> str:
    if not path:
        return ""
    try:
        return path.resolve().as_uri()
    except ValueError:
        return ""


def _validate_date(name: str, value: str | None) -> None:
    if value and not DATE_RE.match(value):
        raise ValueError(f"{name} must use YYYYMMDD, for example 20260530.")


def _append_date_args(
    command: list[str],
    from_date: str | None,
    until_date: str | None,
    end_date: str | None,
    days: int,
) -> None:
    if from_date or until_date:
        if not from_date or not until_date:
            raise ValueError("Pass both from_date and until_date, or neither.")
        _validate_date("from_date", from_date)
        _validate_date("until_date", until_date)
        command.extend(["-FromDate", from_date, "-UntilDate", until_date])
        return

    if end_date:
        _validate_date("end_date", end_date)
        command.extend(["-EndDate", end_date])

    if days < 1:
        raise ValueError("days must be at least 1.")
    command.extend(["-Days", str(days)])


def _append_compare_date_args(
    command: list[str],
    from_date: str | None,
    until_date: str | None,
    end_date: str | None,
    days: int,
) -> None:
    if from_date or until_date:
        if not from_date or not until_date:
            raise ValueError("Pass both from_date and until_date, or neither.")
        _validate_date("from_date", from_date)
        _validate_date("until_date", until_date)
        command.extend(["--from-date", from_date, "--until-date", until_date])
        return

    if end_date:
        _validate_date("end_date", end_date)
        command.extend(["--end-date", end_date])

    if days < 1:
        raise ValueError("days must be at least 1.")
    command.extend(["--days", str(days)])


def _redact(text: str) -> str:
    return SECRET_RE.sub(r"\1\2***", text)


def _tail(text: str, limit: int = 4000) -> str:
    text = _redact(text or "")
    if len(text) <= limit:
        return text
    return text[-limit:]


def _run_process(command: list[str], timeout_seconds: int) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        cwd=ROOT,
        text=True,
        capture_output=True,
        stdin=subprocess.DEVNULL,
        timeout=timeout_seconds,
        check=False,
    )


def _powershell_command(script_name: str) -> list[str]:
    powershell = os.environ.get("DGE_POWERSHELL", "powershell")
    return [
        powershell,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(ROOT / script_name),
    ]


def _read_json(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(data, list):
        return [row for row in data if isinstance(row, dict)]
    return []


def _status_counts(summary_rows: list[dict[str, Any]]) -> dict[str, int]:
    return dict(Counter(str(row.get("status", "")) for row in summary_rows))


def _total_difference_count(summary_rows: list[dict[str, Any]]) -> int:
    total = 0
    for row in summary_rows:
        try:
            total += int(row.get("difference_count") or 0)
        except (TypeError, ValueError):
            pass
    return total


def _extract_path(pattern: re.Pattern[str], text: str) -> Path | None:
    match = pattern.search(text or "")
    if not match:
        return None
    return Path(match.group(1).strip())


def _run_suffix_from_dashboard(path: Path | None) -> str | None:
    if not path or not path.name.startswith("dashboard_") or path.suffix.lower() != ".html":
        return None
    return path.stem.removeprefix("dashboard_")


def _latest_run_dashboard(output_dir: Path) -> Path | None:
    dashboards = [
        path
        for path in output_dir.glob("dashboard_*.html")
        if path.name.lower() != "dashboard.html"
    ]
    if not dashboards:
        return None
    return max(dashboards, key=lambda path: path.stat().st_mtime)


def _run_artifacts(output_dir: Path, suffix: str | None) -> dict[str, Path | None]:
    latest_dashboard = output_dir / "dashboard.html"
    latest_summary = output_dir / "summary.json"
    latest_details = output_dir / "details.json"

    if suffix:
        return {
            "dashboard": output_dir / f"dashboard_{suffix}.html",
            "summary": output_dir / f"summary_{suffix}.json",
            "details": output_dir / f"details_{suffix}.json",
            "latest_dashboard": latest_dashboard,
            "latest_summary": latest_summary,
            "latest_details": latest_details,
        }

    return {
        "dashboard": _latest_run_dashboard(output_dir),
        "summary": latest_summary,
        "details": latest_details,
        "latest_dashboard": latest_dashboard,
        "latest_summary": latest_summary,
        "latest_details": latest_details,
    }


def _summarize_run(
    output_dir: Path,
    completed: subprocess.CompletedProcess[str] | None = None,
    details_limit: int = 20,
) -> dict[str, Any]:
    stdout = completed.stdout if completed else ""
    stderr = completed.stderr if completed else ""
    exit_code = completed.returncode if completed else None
    comparison_exit = exit_code in (0, 2, None)
    extracted_run_dashboard = _extract_path(RUN_DASHBOARD_RE, stdout)
    run_dashboard = extracted_run_dashboard or (_latest_run_dashboard(output_dir) if comparison_exit else None)
    latest_dashboard = _extract_path(LATEST_DASHBOARD_RE, stdout) or (output_dir / "dashboard.html")
    suffix = _run_suffix_from_dashboard(run_dashboard)
    artifacts = _run_artifacts(output_dir, suffix)

    summary_path = artifacts["summary"] or (output_dir / "summary.json")
    details_path = artifacts["details"] or (output_dir / "details.json")
    should_read_artifacts = comparison_exit or extracted_run_dashboard is not None
    summary_rows = _read_json(summary_path) if summary_path and should_read_artifacts else []
    detail_rows = _read_json(details_path) if details_path and should_read_artifacts else []
    counts = _status_counts(summary_rows)
    green = bool(summary_rows) and set(counts) == {"MATCH"}

    completed_comparison = comparison_exit and bool(summary_rows)

    return {
        "completed": completed_comparison,
        "green": green,
        "exit_code": exit_code,
        "status_counts": counts,
        "total_difference_count": _total_difference_count(summary_rows),
        "summary_rows": summary_rows,
        "details_preview": detail_rows[: max(0, details_limit)],
        "details_preview_count": min(len(detail_rows), max(0, details_limit)),
        "details_total_count": len(detail_rows),
        "dashboard_path": str(latest_dashboard),
        "dashboard_url": _file_uri(latest_dashboard),
        "run_dashboard_path": str(run_dashboard) if run_dashboard else "",
        "run_dashboard_url": _file_uri(run_dashboard),
        "summary_path": str(summary_path) if summary_path else "",
        "summary_url": _file_uri(summary_path),
        "details_path": str(details_path) if details_path else "",
        "details_url": _file_uri(details_path),
        "stdout_tail": _tail(stdout),
        "stderr_tail": _tail(stderr),
    }


def _resolve_brand_paths(
    config: BrandConfig,
    env_values: dict[str, str],
    sftp_dir: str | None,
    fabric_dir: str | None,
    output_dir: str | None,
    use_generic_overrides: bool,
) -> tuple[Path, Path, Path]:
    resolved_sftp_dir = _resolve_path(
        sftp_dir if use_generic_overrides else None,
        _brand_env(env_values, config, "SFTP_LOCAL_DIR"),
        _brand_env(env_values, config, "BLOB_LOCAL_DIR"),
        _generic_env(env_values, "DGE_SFTP_LOCAL_DIR") if use_generic_overrides else None,
        _generic_env(env_values, "DGE_BLOB_LOCAL_DIR") if use_generic_overrides else None,
        config.default_sftp_dir,
    )
    resolved_fabric_dir = _resolve_path(
        fabric_dir if use_generic_overrides else None,
        _brand_env(env_values, config, "FABRIC_DIR"),
        _generic_env(env_values, "DGE_FABRIC_DIR") if use_generic_overrides else None,
        config.default_fabric_dir,
    )
    resolved_output_dir = _resolve_path(
        output_dir if use_generic_overrides else None,
        _brand_env(env_values, config, "OUTPUT_DIR"),
        _generic_env(env_values, "DGE_OUTPUT_DIR") if use_generic_overrides else None,
        config.default_output_dir,
    )
    return resolved_sftp_dir, resolved_fabric_dir, resolved_output_dir


def _run_brand_reconciliation(
    config: BrandConfig,
    source_mode: str = "local",
    env_file: str | None = None,
    sftp_dir: str | None = None,
    fabric_dir: str | None = None,
    output_dir: str | None = None,
    storage_prefix: str | None = None,
    sftp_remote_dir: str | None = None,
    fabric_folder: str | None = None,
    from_date: str | None = None,
    until_date: str | None = None,
    end_date: str | None = None,
    days: int = 3,
    force_download: bool = False,
    skip_fabric_download: bool = False,
    python_path: str | None = None,
    timeout_seconds: int = 1800,
    use_generic_overrides: bool = True,
) -> dict[str, Any]:
    source_mode = source_mode.lower().strip()
    env_values = _read_dotenv(env_file)

    resolved_sftp_dir, resolved_fabric_dir, resolved_output_dir = _resolve_brand_paths(
        config=config,
        env_values=env_values,
        sftp_dir=sftp_dir,
        fabric_dir=fabric_dir,
        output_dir=output_dir,
        use_generic_overrides=use_generic_overrides,
    )
    resolved_python = python_path or sys.executable
    resolved_storage_prefix = _first_value(
        storage_prefix if use_generic_overrides else None,
        _brand_env(env_values, config, "STORAGE_PREFIX"),
        _generic_env(env_values, "DGE_STORAGE_PREFIX") if use_generic_overrides else None,
        config.default_storage_prefix,
    )
    resolved_sftp_remote_dir = _first_value(
        sftp_remote_dir if use_generic_overrides else None,
        _brand_env(env_values, config, "SFTP_REMOTE_DIR"),
        _generic_env(env_values, "DGE_SFTP_REMOTE_DIR") if use_generic_overrides else None,
        config.default_sftp_remote_dir,
    )
    resolved_fabric_folder = _first_value(
        fabric_folder if use_generic_overrides else None,
        _brand_env(env_values, config, "FABRIC_FOLDER"),
        _generic_env(env_values, "FABRIC_FOLDER") if use_generic_overrides else None,
        config.default_fabric_folder,
    )

    if source_mode == "local":
        command = [
            resolved_python,
            str(ROOT / "dge_compare.py"),
            "--sftp-dir",
            str(resolved_sftp_dir),
            "--fabric-dir",
            str(resolved_fabric_dir),
            "--output-dir",
            str(resolved_output_dir),
            "--brand",
            config.compare_brand,
            "--sftp-source-url",
            STORAGE_SOURCE_URL,
            "--fabric-source-url",
            _fabric_source_url(resolved_fabric_folder),
        ]
        brand_aliases = []
        if config.sftp_brand != config.compare_brand:
            brand_aliases.append(f"{config.sftp_brand}={config.compare_brand}")
        if config.fabric_brand != config.compare_brand:
            brand_aliases.append(f"{config.fabric_brand}={config.compare_brand}")
        for alias in sorted(set(brand_aliases)):
            command.extend(["--brand-alias", alias])
        for report in DGE_REPORTS:
            command.extend(["--report", report])
        _append_compare_date_args(command, from_date, until_date, end_date, days)
    elif source_mode == "blob":
        command = _powershell_command("run_caesars_blob_local.ps1")
        if env_file:
            command.extend(["-EnvFile", env_file])
        if resolved_storage_prefix:
            command.extend(["-Prefix", resolved_storage_prefix])
        command.extend(
            [
                "-Brand",
                config.compare_brand,
                "-SftpBrand",
                config.sftp_brand,
                "-FabricBrand",
                config.fabric_brand,
                "-FabricFolder",
                resolved_fabric_folder,
                "-BlobLocalDir",
                str(resolved_sftp_dir),
                "-FabricDir",
                str(resolved_fabric_dir),
                "-OutputDir",
                str(resolved_output_dir),
                "-Python",
                resolved_python,
            ]
        )
        _append_date_args(command, from_date, until_date, end_date, days)
    elif source_mode == "sftp":
        command = _powershell_command("run_caesars_sftp_local.ps1")
        if env_file:
            command.extend(["-EnvFile", env_file])
        if resolved_sftp_remote_dir:
            command.extend(["-SftpRemoteDir", resolved_sftp_remote_dir])
        command.extend(
            [
                "-Brand",
                config.compare_brand,
                "-SftpBrand",
                config.sftp_brand,
                "-FabricBrand",
                config.fabric_brand,
                "-FabricFolder",
                resolved_fabric_folder,
                "-SftpLocalDir",
                str(resolved_sftp_dir),
                "-FabricDir",
                str(resolved_fabric_dir),
                "-OutputDir",
                str(resolved_output_dir),
                "-Python",
                resolved_python,
            ]
        )
        _append_date_args(command, from_date, until_date, end_date, days)
    else:
        raise ValueError("source_mode must be local, blob, or sftp.")

    if force_download and source_mode in {"blob", "sftp"}:
        command.append("-ForceDownload")
    if skip_fabric_download and source_mode in {"blob", "sftp"}:
        command.append("-SkipFabricDownload")

    completed = _run_process(command, timeout_seconds)
    result = _summarize_run(resolved_output_dir, completed)
    result["brand"] = config.key
    result["compare_brand"] = config.compare_brand
    result["sftp_brand"] = config.sftp_brand
    result["fabric_brand"] = config.fabric_brand
    result["fabric_folder"] = resolved_fabric_folder
    result["storage_prefix"] = resolved_storage_prefix or ""
    result["sftp_remote_dir"] = resolved_sftp_remote_dir or ""
    result["source_mode"] = source_mode
    result["sftp_dir"] = str(resolved_sftp_dir)
    result["fabric_dir"] = str(resolved_fabric_dir)
    result["output_dir"] = str(resolved_output_dir)
    result["python_path"] = resolved_python
    result["command_status"] = "comparison_completed" if result["completed"] else "failed_before_summary"
    return result


def _aggregate_brand_results(results: list[dict[str, Any]]) -> dict[str, Any]:
    if len(results) == 1:
        return results[0]

    status_counts: Counter[str] = Counter()
    summary_rows: list[dict[str, Any]] = []
    details_preview: list[dict[str, Any]] = []
    total_difference_count = 0
    details_total_count = 0
    for result in results:
        status_counts.update(result.get("status_counts", {}))
        summary_rows.extend(result.get("summary_rows", []))
        details_preview.extend(result.get("details_preview", []))
        total_difference_count += int(result.get("total_difference_count") or 0)
        details_total_count += int(result.get("details_total_count") or 0)

    return {
        "completed": all(bool(result.get("completed")) for result in results),
        "green": all(bool(result.get("green")) for result in results),
        "exit_code": 0 if all(int(result.get("exit_code") or 0) == 0 for result in results) else 2,
        "status_counts": dict(status_counts),
        "total_difference_count": total_difference_count,
        "summary_rows": summary_rows,
        "details_preview": details_preview[:20],
        "details_preview_count": min(len(details_preview), 20),
        "details_total_count": details_total_count,
        "brands": results,
        "source_mode": results[0].get("source_mode") if results else "",
        "command_status": "comparison_completed" if all(bool(result.get("completed")) for result in results) else "failed_before_summary",
    }


def _run_dge_reconciliation_impl(
    brand: str = "caesars",
    source_mode: str = "local",
    env_file: str | None = None,
    sftp_dir: str | None = None,
    fabric_dir: str | None = None,
    output_dir: str | None = None,
    storage_prefix: str | None = None,
    sftp_remote_dir: str | None = None,
    fabric_folder: str | None = None,
    from_date: str | None = None,
    until_date: str | None = None,
    end_date: str | None = None,
    days: int = 3,
    force_download: bool = False,
    skip_fabric_download: bool = False,
    python_path: str | None = None,
    timeout_seconds: int = 1800,
) -> dict[str, Any]:
    configs = _select_brand_configs(brand)
    use_generic_overrides = len(configs) == 1
    results = [
        _run_brand_reconciliation(
            config=config,
            source_mode=source_mode,
            env_file=env_file,
            sftp_dir=sftp_dir,
            fabric_dir=fabric_dir,
            output_dir=output_dir,
            storage_prefix=storage_prefix,
            sftp_remote_dir=sftp_remote_dir,
            fabric_folder=fabric_folder,
            from_date=from_date,
            until_date=until_date,
            end_date=end_date,
            days=days,
            force_download=force_download,
            skip_fabric_download=skip_fabric_download,
            python_path=python_path,
            timeout_seconds=timeout_seconds,
            use_generic_overrides=use_generic_overrides,
        )
        for config in configs
    ]
    return _aggregate_brand_results(results)


@mcp.tool()
def run_dge_reconciliation(
    brand: str = "caesars",
    source_mode: str = "local",
    env_file: str | None = None,
    sftp_dir: str | None = None,
    fabric_dir: str | None = None,
    output_dir: str | None = None,
    storage_prefix: str | None = None,
    sftp_remote_dir: str | None = None,
    fabric_folder: str | None = None,
    from_date: str | None = None,
    until_date: str | None = None,
    end_date: str | None = None,
    days: int = 3,
    force_download: bool = False,
    skip_fabric_download: bool = False,
    python_path: str | None = None,
    timeout_seconds: int = 1800,
) -> dict[str, Any]:
    """Run DGE reconciliation for brand=caesars, brand=fd, or brand=all."""

    return _run_dge_reconciliation_impl(
        brand=brand,
        source_mode=source_mode,
        env_file=env_file,
        sftp_dir=sftp_dir,
        fabric_dir=fabric_dir,
        output_dir=output_dir,
        storage_prefix=storage_prefix,
        sftp_remote_dir=sftp_remote_dir,
        fabric_folder=fabric_folder,
        from_date=from_date,
        until_date=until_date,
        end_date=end_date,
        days=days,
        force_download=force_download,
        skip_fabric_download=skip_fabric_download,
        python_path=python_path,
        timeout_seconds=timeout_seconds,
    )


@mcp.tool()
def run_caesars_reconciliation(
    source_mode: str = "local",
    env_file: str | None = None,
    sftp_dir: str | None = None,
    fabric_dir: str | None = None,
    output_dir: str | None = None,
    from_date: str | None = None,
    until_date: str | None = None,
    end_date: str | None = None,
    days: int = 3,
    force_download: bool = False,
    skip_fabric_download: bool = False,
    python_path: str | None = None,
    timeout_seconds: int = 1800,
) -> dict[str, Any]:
    """Run the Caesars DGE reconciliation. Kept for backward compatibility."""

    return _run_dge_reconciliation_impl(
        brand="caesars",
        source_mode=source_mode,
        env_file=env_file,
        sftp_dir=sftp_dir,
        fabric_dir=fabric_dir,
        output_dir=output_dir,
        from_date=from_date,
        until_date=until_date,
        end_date=end_date,
        days=days,
        force_download=force_download,
        skip_fabric_download=skip_fabric_download,
        python_path=python_path,
        timeout_seconds=timeout_seconds,
    )


@mcp.tool()
def list_report_runs(output_dir: str = str(DEFAULT_OUTPUT_DIR), limit: int = 10) -> dict[str, Any]:
    """List archived reconciliation dashboard runs from an output folder."""

    output = Path(output_dir).expanduser()
    if not output.exists():
        return {"output_dir": str(output), "runs": []}

    dashboards = [
        path
        for path in output.glob("dashboard_*.html")
        if path.name.lower() != "dashboard.html"
    ]
    dashboards.sort(key=lambda path: path.stat().st_mtime, reverse=True)

    runs: list[dict[str, Any]] = []
    for dashboard in dashboards[: max(0, limit)]:
        suffix = _run_suffix_from_dashboard(dashboard)
        artifacts = _run_artifacts(output, suffix)
        summary_path = artifacts["summary"]
        details_path = artifacts["details"]
        summary_rows = _read_json(summary_path) if summary_path else []
        counts = _status_counts(summary_rows)
        runs.append(
            {
                "suffix": suffix,
                "modified_at": datetime.fromtimestamp(
                    dashboard.stat().st_mtime,
                    tz=timezone.utc,
                ).isoformat(),
                "green": bool(summary_rows) and set(counts) == {"MATCH"},
                "status_counts": counts,
                "total_difference_count": _total_difference_count(summary_rows),
                "dashboard_path": str(dashboard),
                "dashboard_url": _file_uri(dashboard),
                "summary_path": str(summary_path) if summary_path else "",
                "summary_url": _file_uri(summary_path),
                "details_path": str(details_path) if details_path else "",
                "details_url": _file_uri(details_path),
            }
        )

    return {
        "output_dir": str(output),
        "latest_dashboard_path": str(output / "dashboard.html"),
        "latest_dashboard_url": _file_uri(output / "dashboard.html"),
        "runs": runs,
    }


@mcp.tool()
def get_report_run(
    output_dir: str = str(DEFAULT_OUTPUT_DIR),
    suffix: str | None = None,
    max_details: int = 100,
) -> dict[str, Any]:
    """Read one reconciliation run summary and a limited list of diffs.

    Pass a suffix returned by list_report_runs. When suffix is omitted, this
    returns the newest archived run.
    """

    output = Path(output_dir).expanduser()
    if not output.exists():
        return {"output_dir": str(output), "error": "Output folder does not exist."}

    if not suffix:
        newest = _latest_run_dashboard(output)
        suffix = _run_suffix_from_dashboard(newest)

    artifacts = _run_artifacts(output, suffix)
    summary_path = artifacts["summary"]
    details_path = artifacts["details"]
    summary_rows = _read_json(summary_path) if summary_path else []
    detail_rows = _read_json(details_path) if details_path else []
    counts = _status_counts(summary_rows)

    return {
        "output_dir": str(output),
        "suffix": suffix,
        "green": bool(summary_rows) and set(counts) == {"MATCH"},
        "status_counts": counts,
        "total_difference_count": _total_difference_count(summary_rows),
        "summary_rows": summary_rows,
        "details": detail_rows[: max(0, max_details)],
        "details_returned": min(len(detail_rows), max(0, max_details)),
        "details_total_count": len(detail_rows),
        "dashboard_path": str(artifacts["dashboard"]) if artifacts["dashboard"] else "",
        "dashboard_url": _file_uri(artifacts["dashboard"]),
        "summary_path": str(summary_path) if summary_path else "",
        "summary_url": _file_uri(summary_path),
        "details_path": str(details_path) if details_path else "",
        "details_url": _file_uri(details_path),
    }


if __name__ == "__main__":
    mcp.run()
