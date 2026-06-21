param(
  [string]$EnvFile,

  [string]$SftpHost = $env:DGE_SFTP_HOST,
  [int]$SftpPort = $(if ($env:DGE_SFTP_PORT) { [int]$env:DGE_SFTP_PORT } else { 22 }),
  [string]$SftpUsername = $env:DGE_SFTP_USERNAME,
  [string]$SftpPrivateKey = $env:DGE_SFTP_PRIVATE_KEY,
  [string]$SftpRemoteDir = $env:DGE_SFTP_REMOTE_DIR,
  [string]$Brand = $(if ($env:DGE_BRAND) { $env:DGE_BRAND } else { "caesars-nj" }),
  [string]$SftpBrand = $env:DGE_SFTP_BRAND,
  [string]$FabricBrand = $env:DGE_FABRIC_BRAND,

  [string]$FabricTenantId = $(if ($env:FABRIC_TENANT_ID) { $env:FABRIC_TENANT_ID } else { $env:AZURE_TENANT_ID }),
  [string]$FabricClientId = $(if ($env:FABRIC_CLIENT_ID) { $env:FABRIC_CLIENT_ID } else { $env:AZURE_CLIENT_ID }),
  [string]$FabricClientSecret = $(if ($env:FABRIC_CLIENT_SECRET) { $env:FABRIC_CLIENT_SECRET } else { $env:AZURE_CLIENT_SECRET }),
  [string]$FabricWorkspaceId = $env:FABRIC_WORKSPACE_ID,
  [string]$FabricLakehouseId = $env:FABRIC_LAKEHOUSE_ID,
  [string]$FabricFolder = $env:FABRIC_FOLDER,

  [string]$SftpLocalDir = $env:DGE_SFTP_LOCAL_DIR,
  [string]$FabricDir = $env:DGE_FABRIC_DIR,
  [string]$OutputDir = $env:DGE_OUTPUT_DIR,

  [int]$Days = 3,
  [string]$FromDate,
  [string]$UntilDate,
  [string]$EndDate,
  [string]$Python = "python",
  [Alias("OverwriteDownloadedFiles")]
  [switch]$ForceDownload,
  [switch]$SkipFabricDownload,
  [switch]$OpenDashboard
)

$ErrorActionPreference = "Stop"

function Import-DotEnv {
  param([string]$Path)

  if (-not $Path) {
    return
  }
  if (-not (Test-Path -LiteralPath $Path)) {
    throw "Env file not found: $Path"
  }

  Get-Content -LiteralPath $Path | ForEach-Object {
    $line = $_.Trim()
    if (-not $line -or $line.StartsWith("#")) {
      return
    }
    $separator = $line.IndexOf("=")
    if ($separator -le 0) {
      return
    }
    $name = $line.Substring(0, $separator).Trim()
    $value = $line.Substring($separator + 1).Trim().Trim('"')
    [Environment]::SetEnvironmentVariable($name, $value, "Process")
  }
}

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

function Resolve-BrandSlug {
  param([string]$Value)

  switch ($Value.ToLowerInvariant()) {
    "caesars" { return "caesars" }
    "caesars-nj" { return "caesars" }
    "fd" { return "fd" }
    "fanduel" { return "fd" }
    "fanduel-nj" { return "fd" }
    default { return $Value.ToLowerInvariant().Replace("-nj", "") }
  }
}

Import-DotEnv -Path $EnvFile

if (-not $PSBoundParameters.ContainsKey("Brand") -and $env:DGE_BRAND) { $Brand = $env:DGE_BRAND }
$compareBrand = Resolve-CompareBrand -Value $Brand
if (-not $SftpBrand) {
  if ($env:DGE_SFTP_BRAND) { $SftpBrand = $env:DGE_SFTP_BRAND } else { $SftpBrand = Resolve-SftpBrand -Value $Brand }
}
if (-not $FabricBrand) {
  if ($env:DGE_FABRIC_BRAND) { $FabricBrand = $env:DGE_FABRIC_BRAND } else { $FabricBrand = Resolve-FabricBrand -Value $Brand }
}
$brandSlug = Resolve-BrandSlug -Value $Brand

if (-not $SftpHost) { $SftpHost = $env:DGE_SFTP_HOST }
if (-not $PSBoundParameters.ContainsKey("SftpPort") -and $env:DGE_SFTP_PORT) { $SftpPort = [int]$env:DGE_SFTP_PORT }
if (-not $SftpUsername) { $SftpUsername = $env:DGE_SFTP_USERNAME }
if (-not $SftpPrivateKey) { $SftpPrivateKey = $env:DGE_SFTP_PRIVATE_KEY }
if (-not $SftpRemoteDir) { $SftpRemoteDir = $env:DGE_SFTP_REMOTE_DIR }
if (-not $SftpRemoteDir -and $brandSlug -eq "fd") { $SftpRemoteDir = "/c8-nj-prod/fd" }
if (-not $FabricTenantId) {
  if ($env:FABRIC_TENANT_ID) { $FabricTenantId = $env:FABRIC_TENANT_ID } else { $FabricTenantId = $env:AZURE_TENANT_ID }
}
if (-not $FabricClientId) {
  if ($env:FABRIC_CLIENT_ID) { $FabricClientId = $env:FABRIC_CLIENT_ID } else { $FabricClientId = $env:AZURE_CLIENT_ID }
}
if (-not $FabricClientSecret) {
  if ($env:FABRIC_CLIENT_SECRET) { $FabricClientSecret = $env:FABRIC_CLIENT_SECRET } else { $FabricClientSecret = $env:AZURE_CLIENT_SECRET }
}
if (-not $FabricWorkspaceId) { $FabricWorkspaceId = $env:FABRIC_WORKSPACE_ID }
if (-not $FabricLakehouseId) { $FabricLakehouseId = $env:FABRIC_LAKEHOUSE_ID }
if (-not $FabricFolder) { $FabricFolder = $env:FABRIC_FOLDER }
if (-not $FabricFolder) { $FabricFolder = "Files/auto_reports/regulatory/$FabricBrand" }
if (-not $FabricDir) { $FabricDir = $env:DGE_FABRIC_DIR }
if (-not $FabricDir) { $FabricDir = "C:\tmp\dge\$brandSlug\fabric" }
if (-not $PSBoundParameters.ContainsKey("SftpLocalDir") -and $env:DGE_SFTP_LOCAL_DIR) { $SftpLocalDir = $env:DGE_SFTP_LOCAL_DIR }
if (-not $SftpLocalDir) { $SftpLocalDir = "C:\tmp\dge\$brandSlug\sftp" }
if (-not $PSBoundParameters.ContainsKey("OutputDir") -and $env:DGE_OUTPUT_DIR) { $OutputDir = $env:DGE_OUTPUT_DIR }
if (-not $OutputDir) { $OutputDir = "C:\tmp\dge\$brandSlug\out" }

$required = @{
  "DGE_SFTP_HOST" = $SftpHost
  "DGE_SFTP_USERNAME" = $SftpUsername
  "DGE_SFTP_PRIVATE_KEY" = $SftpPrivateKey
  "DGE_SFTP_REMOTE_DIR" = $SftpRemoteDir
  "DGE_FABRIC_DIR" = $FabricDir
}

foreach ($item in $required.GetEnumerator()) {
  if (-not $item.Value) {
    throw "Missing required value: $($item.Key). Pass it as a parameter or set it in -EnvFile."
  }
}

$downloadScript = Join-Path $PSScriptRoot "download_caesars_sftp.ps1"
$fabricDownloadScript = Join-Path $PSScriptRoot "download_caesars_fabric.ps1"
$compareScript = Join-Path $PSScriptRoot "run_caesars_local.ps1"

$downloadParams = @{
  HostName = $SftpHost
  Port = $SftpPort
  Username = $SftpUsername
  PrivateKey = $SftpPrivateKey
  RemoteDir = $SftpRemoteDir
  LocalDir = $SftpLocalDir
  Brand = $SftpBrand
  Days = $Days
}

if ($FromDate -or $UntilDate) {
  if (-not $FromDate -or -not $UntilDate) {
    throw "Pass both -FromDate and -UntilDate, or neither."
  }
  $downloadParams.FromDate = $FromDate
  $downloadParams.UntilDate = $UntilDate
} elseif ($EndDate) {
  $downloadParams.EndDate = $EndDate
}
if ($ForceDownload) {
  $downloadParams.ForceDownload = $true
}

& $downloadScript @downloadParams
if (-not $?) {
  exit 1
}

$fabricDownloadConfigured = $FabricWorkspaceId -and $FabricLakehouseId -and $FabricFolder
if ($fabricDownloadConfigured -and -not $SkipFabricDownload) {
  $fabricDownloadParams = @{
    TenantId = $FabricTenantId
    ClientId = $FabricClientId
    ClientSecret = $FabricClientSecret
    WorkspaceId = $FabricWorkspaceId
    LakehouseId = $FabricLakehouseId
    Folder = $FabricFolder
    LocalDir = $FabricDir
    Brand = $FabricBrand
    Days = $Days
  }

  if ($FromDate -or $UntilDate) {
    $fabricDownloadParams.FromDate = $FromDate
    $fabricDownloadParams.UntilDate = $UntilDate
  } elseif ($EndDate) {
    $fabricDownloadParams.EndDate = $EndDate
  }
  if ($ForceDownload) {
    $fabricDownloadParams.ForceDownload = $true
  }

  & $fabricDownloadScript @fabricDownloadParams
  if (-not $?) {
    exit 1
  }
} else {
  Write-Host "Fabric OneLake download is not configured or was skipped. Using local Fabric folder: $FabricDir"
}

$compareParams = @{
  SftpDir = $SftpLocalDir
  FabricDir = $FabricDir
  OutputDir = $OutputDir
  Brand = $compareBrand
  SftpBrand = $SftpBrand
  FabricBrand = $FabricBrand
  Days = $Days
  Python = $Python
}

if ($FromDate -or $UntilDate) {
  if (-not $FromDate -or -not $UntilDate) {
    throw "Pass both -FromDate and -UntilDate, or neither."
  }
  $compareParams.FromDate = $FromDate
  $compareParams.UntilDate = $UntilDate
} elseif ($EndDate) {
  $compareParams.EndDate = $EndDate
}
if ($OpenDashboard) {
  $compareParams.OpenDashboard = $true
}

& $compareScript @compareParams
exit $LASTEXITCODE
