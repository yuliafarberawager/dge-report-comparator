param(
  [Parameter(Mandatory = $true)]
  [string]$SftpDir,

  [Parameter(Mandatory = $true)]
  [string]$FabricDir,

  [string]$OutputDir = (Join-Path $PSScriptRoot "out-caesars"),
  [string]$Brand = $(if ($env:DGE_BRAND) { $env:DGE_BRAND } else { "caesars-nj" }),
  [string]$SftpBrand = $env:DGE_SFTP_BRAND,
  [string]$FabricBrand = $env:DGE_FABRIC_BRAND,
  [string[]]$BrandAlias = @(),
  [int]$Days = 2,
  [string]$FromDate,
  [string]$UntilDate,
  [string]$EndDate,
  [string]$Python = "python",
  [switch]$OpenDashboard
)

function Resolve-CompareBrand {
  param([string]$Value)

  switch ($Value.ToLowerInvariant()) {
    "caesars" { return "caesars-nj" }
    "caesars-nj" { return "caesars-nj" }
    "fd" { return "fd" }
    "fanduel" { return "fd" }
    "fanduel-nj" { return "fd" }
    default { return $Value }
  }
}

function Resolve-SftpBrand {
  param([string]$Value)

  switch ($Value.ToLowerInvariant()) {
    "fd" { return "fd" }
    "fanduel" { return "fd" }
    "fanduel-nj" { return "fd" }
    default { return $Value }
  }
}

function Resolve-FabricBrand {
  param([string]$Value)

  switch ($Value.ToLowerInvariant()) {
    "fd" { return "fanduel-nj" }
    "fanduel" { return "fanduel-nj" }
    "fanduel-nj" { return "fanduel-nj" }
    default { return $Value }
  }
}

$compareBrand = Resolve-CompareBrand -Value $Brand
if (-not $SftpBrand) { $SftpBrand = Resolve-SftpBrand -Value $Brand }
if (-not $FabricBrand) { $FabricBrand = Resolve-FabricBrand -Value $Brand }

$fabricSourcePath = [System.Uri]::EscapeDataString("Files/auto_reports/regulatory/$FabricBrand")
$fabricSourceUrl = "https://app.fabric.microsoft.com/groups/912d99dd-0947-4c54-ab3b-166a11cf0f0e/lakehouses/f4d2b03c-4a71-42ac-887b-6394865db591?experience=power-bi&selectedPath=$fabricSourcePath&extensionScenario=openArtifact"
$sftpSourceUrl = "https://portal.azure.com/#view/Microsoft_Azure_Storage/ContainerMenuBlade/~/overview/storageAccountId/%2Fsubscriptions%2Ffcb6ca50-8dc2-4c83-b963-1734f7f053ec%2FresourceGroups%2Fawager%2Fproviders%2FMicrosoft.Storage%2FstorageAccounts%2Fawagersftp/path/c8-nj-prod/etag/%220x8DD11EBF1A834F9%22/defaultId//publicAccessVal/None"

$compareScript = Join-Path $PSScriptRoot "dge_compare.py"
$reports = @(
  "game_summary_report",
  "machine_summary_report",
  "pending_transaction",
  "void_transaction"
)

$compareArgs = @(
  $compareScript,
  "--sftp-dir", $SftpDir,
  "--fabric-dir", $FabricDir,
  "--output-dir", $OutputDir,
  "--brand", $compareBrand,
  "--days", $Days,
  "--sftp-source-url", $sftpSourceUrl,
  "--fabric-source-url", $fabricSourceUrl
)

foreach ($report in $reports) {
  $compareArgs += @("--report", $report)
}

$aliases = @($BrandAlias)
if ($SftpBrand -and $SftpBrand -ne $compareBrand) {
  $aliases += "$SftpBrand=$compareBrand"
}
if ($FabricBrand -and $FabricBrand -ne $compareBrand) {
  $aliases += "$FabricBrand=$compareBrand"
}
foreach ($alias in ($aliases | Select-Object -Unique)) {
  $compareArgs += @("--brand-alias", $alias)
}

if ($FromDate -or $UntilDate) {
  if (-not $FromDate -or -not $UntilDate) {
    throw "Pass both -FromDate and -UntilDate, or neither."
  }
  $compareArgs += @("--from-date", $FromDate, "--until-date", $UntilDate)
} elseif ($EndDate) {
  $compareArgs += @("--end-date", $EndDate)
}

& $Python @compareArgs
$compareExitCode = $LASTEXITCODE

$dashboardPath = Join-Path $OutputDir "dashboard.html"
if (Test-Path -LiteralPath $dashboardPath) {
  Write-Host "Dashboard: $dashboardPath"
  if ($OpenDashboard) {
    Start-Process -FilePath $dashboardPath
  }
}

exit $compareExitCode
