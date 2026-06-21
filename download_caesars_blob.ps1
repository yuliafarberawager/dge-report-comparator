param(
  [string]$TenantId = $env:AZURE_TENANT_ID,
  [string]$ClientId = $env:AZURE_CLIENT_ID,
  [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,

  [string]$StorageAccount = $(if ($env:DGE_STORAGE_ACCOUNT) { $env:DGE_STORAGE_ACCOUNT } else { "awagersftp" }),
  [string]$Container = $(if ($env:DGE_STORAGE_CONTAINER) { $env:DGE_STORAGE_CONTAINER } else { "c8-nj-prod" }),
  [string]$Prefix = $env:DGE_STORAGE_PREFIX,

  [Parameter(Mandatory = $true)]
  [string]$LocalDir,

  [int]$Days = 3,
  [string]$FromDate,
  [string]$UntilDate,
  [string]$EndDate,
  [string]$Brand = "caesars-nj",
  [string[]]$ReportPrefixes = @(
    "game_summary_report",
    "machine_summary_report",
    "pending_transaction",
    "void_transaction"
  ),
  [Alias("OverwriteDownloadedFiles")]
  [switch]$ForceDownload,
  [string]$AzExecutable = "az"
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command $AzExecutable -ErrorAction SilentlyContinue)) {
  throw "Cannot find '$AzExecutable'. Install Azure CLI or pass -AzExecutable."
}

New-Item -ItemType Directory -Force -Path $LocalDir | Out-Null

if ($TenantId -and $ClientId -and $ClientSecret) {
  Write-Host "Logging in with service principal $ClientId"
  & $AzExecutable login `
    --service-principal `
    --username $ClientId `
    --password="$ClientSecret" `
    --tenant $TenantId `
    --only-show-errors `
    --output none
  if ($LASTEXITCODE -ne 0) {
    throw "Azure CLI service-principal login failed with exit code $LASTEXITCODE"
  }
} else {
  Write-Host "AZURE_TENANT_ID/AZURE_CLIENT_ID/AZURE_CLIENT_SECRET not fully set. Using current Azure CLI login."
}

if ($FromDate -or $UntilDate) {
  if (-not $FromDate -or -not $UntilDate) {
    throw "Pass both -FromDate and -UntilDate, or neither."
  }
  $firstDate = [datetime]::ParseExact($FromDate, "yyyyMMdd", $null)
  $lastDate = [datetime]::ParseExact($UntilDate, "yyyyMMdd", $null)
  if ($firstDate -gt $lastDate) {
    throw "-FromDate must be before or equal to -UntilDate."
  }
  $dateSet = @{}
  for ($date = $firstDate; $date -le $lastDate; $date = $date.AddDays(1)) {
    $dateSet[$date.ToString("yyyyMMdd")] = $true
  }
} elseif ($EndDate) {
  $lastDate = [datetime]::ParseExact($EndDate, "yyyyMMdd", $null)
  $dateSet = @{}
  for ($i = 0; $i -lt $Days; $i++) {
    $dateSet[$lastDate.AddDays(-$i).ToString("yyyyMMdd")] = $true
  }
} else {
  $lastDate = (Get-Date).Date.AddDays(-1)
  $dateSet = @{}
  for ($i = 0; $i -lt $Days; $i++) {
    $dateSet[$lastDate.AddDays(-$i).ToString("yyyyMMdd")] = $true
  }
}

Write-Host "Listing blobs from $StorageAccount/$Container prefix '$Prefix'"
$listArgs = @(
  "storage", "blob", "list",
  "--account-name", $StorageAccount,
  "--container-name", $Container,
  "--auth-mode", "login",
  "--only-show-errors",
  "--output", "json"
)

if ($Prefix) {
  $listArgs += @("--prefix", $Prefix)
}

$json = & $AzExecutable @listArgs
if ($LASTEXITCODE -ne 0) {
  throw "Azure blob list failed with exit code $LASTEXITCODE"
}

$blobs = ($json -join [Environment]::NewLine) | ConvertFrom-Json
$matches = @()
foreach ($blob in $blobs) {
  if (-not $blob.name) {
    continue
  }
  if (-not $blob.name.EndsWith(".xlsx", [System.StringComparison]::OrdinalIgnoreCase)) {
    continue
  }
  if (-not $blob.name.Contains($Brand)) {
    continue
  }
  $normalizedBlobName = $blob.name.ToLowerInvariant().Replace("-", "_")
  $isWantedReport = $false
  foreach ($reportPrefix in $ReportPrefixes) {
    $normalizedReportPrefix = $reportPrefix.ToLowerInvariant().Replace("-", "_")
    if ($normalizedBlobName.Contains($normalizedReportPrefix)) {
      $isWantedReport = $true
      break
    }
  }
  if (-not $isWantedReport) {
    continue
  }
  foreach ($date in $dateSet.Keys) {
    if ($blob.name.Contains($date)) {
      $matches += $blob
      break
    }
  }
}

Write-Host "Matched blobs: $($matches.Count)"
$downloadedCount = 0
$skippedCount = 0
foreach ($blob in $matches) {
  $fileName = [System.IO.Path]::GetFileName($blob.name)
  $destination = Join-Path $LocalDir $fileName
  if ((Test-Path -LiteralPath $destination) -and -not $ForceDownload) {
    Write-Host "Skipping existing file: $destination"
    $skippedCount++
    continue
  }
  Write-Host "Downloading $($blob.name) -> $destination"
  & $AzExecutable storage blob download `
    --account-name $StorageAccount `
    --container-name $Container `
    --name $blob.name `
    --file $destination `
    --auth-mode login `
    --overwrite true `
    --only-show-errors `
    --output none
  if ($LASTEXITCODE -ne 0) {
    throw "Azure blob download failed for $($blob.name) with exit code $LASTEXITCODE"
  }
  $downloadedCount++
}

Write-Host "Downloaded: $downloadedCount; skipped existing: $skippedCount"
Write-Host "Downloaded files in $LocalDir`: $((Get-ChildItem -LiteralPath $LocalDir -Filter '*.xlsx' -File -ErrorAction SilentlyContinue).Count)"
