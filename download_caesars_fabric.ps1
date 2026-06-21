param(
  [string]$EnvFile,

  [string]$TenantId = $(if ($env:FABRIC_TENANT_ID) { $env:FABRIC_TENANT_ID } else { $env:AZURE_TENANT_ID }),
  [string]$ClientId = $(if ($env:FABRIC_CLIENT_ID) { $env:FABRIC_CLIENT_ID } else { $env:AZURE_CLIENT_ID }),
  [string]$ClientSecret = $(if ($env:FABRIC_CLIENT_SECRET) { $env:FABRIC_CLIENT_SECRET } else { $env:AZURE_CLIENT_SECRET }),

  [string]$WorkspaceId = $env:FABRIC_WORKSPACE_ID,
  [string]$LakehouseId = $env:FABRIC_LAKEHOUSE_ID,
  [string]$Folder = $(if ($env:FABRIC_FOLDER) { $env:FABRIC_FOLDER } else { "Files/auto_reports/regulatory/caesars-nj" }),
  [string]$OneLakeHost = $(if ($env:FABRIC_ONELAKE_HOST) { $env:FABRIC_ONELAKE_HOST } else { "onelake.dfs.fabric.microsoft.com" }),
  [string]$TokenScope = $(if ($env:FABRIC_TOKEN_SCOPE) { $env:FABRIC_TOKEN_SCOPE } else { "https://storage.azure.com/.default" }),

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
  [switch]$ForceDownload
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

function Assert-RequiredValue {
  param(
    [string]$Name,
    [string]$Value
  )

  if (-not $Value) {
    throw "Missing required value: $Name. Pass it as a parameter or set it in -EnvFile."
  }
}

function Get-TargetDates {
  if ($FromDate -or $UntilDate) {
    if (-not $FromDate -or -not $UntilDate) {
      throw "Pass both -FromDate and -UntilDate, or neither."
    }
    $firstDate = [datetime]::ParseExact($FromDate, "yyyyMMdd", $null)
    $lastDate = [datetime]::ParseExact($UntilDate, "yyyyMMdd", $null)
    if ($firstDate -gt $lastDate) {
      throw "-FromDate must be before or equal to -UntilDate."
    }
    for ($date = $firstDate; $date -le $lastDate; $date = $date.AddDays(1)) {
      $date.ToString("yyyyMMdd")
    }
    return
  }

  if ($EndDate) {
    $lastDate = [datetime]::ParseExact($EndDate, "yyyyMMdd", $null)
  } else {
    $lastDate = (Get-Date).Date.AddDays(-1)
  }

  for ($i = 0; $i -lt $Days; $i++) {
    $lastDate.AddDays(-$i).ToString("yyyyMMdd")
  }
}

function ConvertTo-EncodedPath {
  param([string]$Path)

  return (($Path.Trim("/") -split "/") | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join "/"
}

function Test-WantedReportPath {
  param(
    [string]$Path,
    [hashtable]$DateSet
  )

  if (-not $Path.EndsWith(".xlsx", [System.StringComparison]::OrdinalIgnoreCase)) {
    return $false
  }
  if ($Path.IndexOf($Brand, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
    return $false
  }

  $normalizedPath = $Path.ToLowerInvariant().Replace("-", "_")
  $isWantedReport = $false
  foreach ($reportPrefix in $ReportPrefixes) {
    $normalizedReportPrefix = $reportPrefix.ToLowerInvariant().Replace("-", "_")
    if ($normalizedPath.Contains($normalizedReportPrefix)) {
      $isWantedReport = $true
      break
    }
  }
  if (-not $isWantedReport) {
    return $false
  }

  foreach ($date in $DateSet.Keys) {
    if ($Path.Contains($date)) {
      return $true
    }
  }
  return $false
}

function Get-OneLakeAccessToken {
  Write-Host "Requesting Fabric/OneLake token for service principal $ClientId"
  $tokenResponse = Invoke-RestMethod `
    -Method Post `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body @{
      grant_type = "client_credentials"
      client_id = $ClientId
      client_secret = $ClientSecret
      scope = $TokenScope
    }

  return $tokenResponse.access_token
}

function Invoke-OneLakeList {
  param(
    [hashtable]$Headers,
    [string]$Directory
  )

  $paths = @()
  $continuation = $null
  do {
    $query = "resource=filesystem&directory=$([System.Uri]::EscapeDataString($Directory))&recursive=true&maxResults=5000"
    if ($continuation) {
      $query += "&continuation=$([System.Uri]::EscapeDataString($continuation))"
    }

    $uri = "https://$OneLakeHost/$WorkspaceId`?$query"
    $response = Invoke-WebRequest -Method Get -Uri $uri -Headers $Headers -UseBasicParsing
    if ($response.Content) {
      $body = $response.Content | ConvertFrom-Json
      if ($body.paths) {
        $paths += @($body.paths)
      }
    }

    $continuation = $null
    $continuationHeader = $response.Headers["x-ms-continuation"]
    if ($continuationHeader) {
      $continuation = @($continuationHeader)[0]
    }
  } while ($continuation)

  return $paths
}

Import-DotEnv -Path $EnvFile

if (-not $PSBoundParameters.ContainsKey("TenantId")) {
  if ($env:FABRIC_TENANT_ID) { $TenantId = $env:FABRIC_TENANT_ID } else { $TenantId = $env:AZURE_TENANT_ID }
}
if (-not $PSBoundParameters.ContainsKey("ClientId")) {
  if ($env:FABRIC_CLIENT_ID) { $ClientId = $env:FABRIC_CLIENT_ID } else { $ClientId = $env:AZURE_CLIENT_ID }
}
if (-not $PSBoundParameters.ContainsKey("ClientSecret")) {
  if ($env:FABRIC_CLIENT_SECRET) { $ClientSecret = $env:FABRIC_CLIENT_SECRET } else { $ClientSecret = $env:AZURE_CLIENT_SECRET }
}
if (-not $PSBoundParameters.ContainsKey("WorkspaceId") -and $env:FABRIC_WORKSPACE_ID) { $WorkspaceId = $env:FABRIC_WORKSPACE_ID }
if (-not $PSBoundParameters.ContainsKey("LakehouseId") -and $env:FABRIC_LAKEHOUSE_ID) { $LakehouseId = $env:FABRIC_LAKEHOUSE_ID }
if (-not $PSBoundParameters.ContainsKey("Folder") -and $env:FABRIC_FOLDER) { $Folder = $env:FABRIC_FOLDER }
if (-not $PSBoundParameters.ContainsKey("OneLakeHost") -and $env:FABRIC_ONELAKE_HOST) { $OneLakeHost = $env:FABRIC_ONELAKE_HOST }
if (-not $PSBoundParameters.ContainsKey("TokenScope") -and $env:FABRIC_TOKEN_SCOPE) { $TokenScope = $env:FABRIC_TOKEN_SCOPE }

Assert-RequiredValue -Name "FABRIC_TENANT_ID or AZURE_TENANT_ID" -Value $TenantId
Assert-RequiredValue -Name "FABRIC_CLIENT_ID or AZURE_CLIENT_ID" -Value $ClientId
Assert-RequiredValue -Name "FABRIC_CLIENT_SECRET or AZURE_CLIENT_SECRET" -Value $ClientSecret
Assert-RequiredValue -Name "FABRIC_WORKSPACE_ID" -Value $WorkspaceId
Assert-RequiredValue -Name "FABRIC_LAKEHOUSE_ID" -Value $LakehouseId
Assert-RequiredValue -Name "FABRIC_FOLDER" -Value $Folder

New-Item -ItemType Directory -Force -Path $LocalDir | Out-Null

$dates = @(Get-TargetDates)
$dateSet = @{}
foreach ($date in $dates) {
  $dateSet[$date] = $true
}

$oneLakeDirectory = "$LakehouseId/$($Folder.Trim('/'))"
$token = Get-OneLakeAccessToken
$headers = @{
  Authorization = "Bearer $token"
  "x-ms-version" = "2021-06-08"
  "x-ms-date" = (Get-Date).ToUniversalTime().ToString("R")
}

Write-Host "Listing Fabric OneLake folder: https://$OneLakeHost/$WorkspaceId/$oneLakeDirectory"
$paths = @(Invoke-OneLakeList -Headers $headers -Directory $oneLakeDirectory)
$files = @($paths | Where-Object { -not $_.isDirectory -and (Test-WantedReportPath -Path $_.name -DateSet $dateSet) })

Write-Host "Matched Fabric files: $($files.Count)"
$downloadedCount = 0
$skippedCount = 0
foreach ($file in $files) {
  $fileName = [System.IO.Path]::GetFileName($file.name)
  $destination = Join-Path $LocalDir $fileName
  if ((Test-Path -LiteralPath $destination) -and -not $ForceDownload) {
    Write-Host "Skipping existing Fabric file: $destination"
    $skippedCount++
    continue
  }

  $encodedPath = ConvertTo-EncodedPath -Path $file.name
  $downloadUri = "https://$OneLakeHost/$WorkspaceId/$encodedPath"
  Write-Host "Downloading Fabric $($file.name) -> $destination"
  Invoke-WebRequest -Method Get -Uri $downloadUri -Headers $headers -OutFile $destination -UseBasicParsing
  $downloadedCount++
}

Write-Host "Fabric downloaded: $downloadedCount; skipped existing: $skippedCount"
Write-Host "Fabric files in $LocalDir`: $((Get-ChildItem -LiteralPath $LocalDir -Filter '*.xlsx' -File -ErrorAction SilentlyContinue).Count)"
