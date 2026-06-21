param(
  [Parameter(Mandatory = $true)]
  [string]$HostName,

  [Parameter(Mandatory = $true)]
  [string]$Username,

  [Parameter(Mandatory = $true)]
  [string]$PrivateKey,

  [Parameter(Mandatory = $true)]
  [string]$RemoteDir,

  [Parameter(Mandatory = $true)]
  [string]$LocalDir,

  [int]$Port = 22,
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
  [string]$SftpExecutable = "sftp"
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command $SftpExecutable -ErrorAction SilentlyContinue)) {
  throw "Cannot find '$SftpExecutable'. Install/use Windows OpenSSH Client or pass -SftpExecutable."
}

if (-not (Test-Path -LiteralPath $PrivateKey)) {
  throw "Private key not found: $PrivateKey"
}

New-Item -ItemType Directory -Force -Path $LocalDir | Out-Null

if ($FromDate -or $UntilDate) {
  if (-not $FromDate -or -not $UntilDate) {
    throw "Pass both -FromDate and -UntilDate, or neither."
  }
  $firstDate = [datetime]::ParseExact($FromDate, "yyyyMMdd", $null)
  $lastDate = [datetime]::ParseExact($UntilDate, "yyyyMMdd", $null)
  if ($firstDate -gt $lastDate) {
    throw "-FromDate must be before or equal to -UntilDate."
  }
  $dates = for ($date = $firstDate; $date -le $lastDate; $date = $date.AddDays(1)) {
    $date.ToString("yyyyMMdd")
  }
} elseif ($EndDate) {
  $lastDate = [datetime]::ParseExact($EndDate, "yyyyMMdd", $null)
  $dates = for ($i = 0; $i -lt $Days; $i++) {
    $lastDate.AddDays(-$i).ToString("yyyyMMdd")
  }
} else {
  $lastDate = (Get-Date).Date.AddDays(-1)
  $dates = for ($i = 0; $i -lt $Days; $i++) {
    $lastDate.AddDays(-$i).ToString("yyyyMMdd")
  }
}

function Get-ReportPatternVariants {
  param([string]$ReportPrefix)

  $variants = @($ReportPrefix)
  if ($ReportPrefix -eq "void_transaction") {
    $variants += "void-transaction"
  }
  foreach ($variant in $variants) {
    $variant
    $variant.ToUpperInvariant()
    $variant.ToLowerInvariant()
  }
}

function Test-ExistingLocalReport {
  param(
    [string]$ReportPrefix,
    [string]$ReportDate
  )

  $normalizedReportPrefix = $ReportPrefix.ToLowerInvariant().Replace("-", "_")
  $existing = Get-ChildItem -LiteralPath $LocalDir -Filter "*$Brand*$ReportDate*.xlsx" -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Name.ToLowerInvariant().Replace("-", "_").Contains($normalizedReportPrefix) } |
    Select-Object -First 1
  return $null -ne $existing
}

$batchFile = New-TemporaryFile
try {
  $commands = New-Object System.Collections.Generic.List[string]
  $commands.Add("cd $RemoteDir")
  $commands.Add("lcd $LocalDir")
  $downloadCommandCount = 0
  $skippedCount = 0
  foreach ($date in $dates) {
    foreach ($reportPrefix in $ReportPrefixes) {
      if ((Test-ExistingLocalReport -ReportPrefix $reportPrefix -ReportDate $date) -and -not $ForceDownload) {
        Write-Host "Skipping existing local $reportPrefix for $Brand $date"
        $skippedCount++
        continue
      }
      $reportPatterns = Get-ReportPatternVariants -ReportPrefix $reportPrefix | Sort-Object -Unique
      foreach ($reportPattern in $reportPatterns) {
        $commands.Add("-mget *$reportPattern*$Brand*$date*.xlsx")
        $downloadCommandCount++
      }
    }
  }
  $commands.Add("bye")

  if ($downloadCommandCount -eq 0) {
    Write-Host "All requested SFTP files already exist locally. Skipping download."
    Write-Host "Skipped existing report/date checks: $skippedCount"
    return
  }

  Set-Content -LiteralPath $batchFile.FullName -Value $commands -Encoding ASCII

  Write-Host "Downloading SFTP files for brand '$Brand' dates: $($dates -join ', ')"
  Write-Host "Remote: $Username@$HostName`:$RemoteDir"
  Write-Host "Local:  $LocalDir"
  if ($ForceDownload) {
    Write-Host "ForceDownload is on: existing local files may be overwritten."
  } elseif ($skippedCount -gt 0) {
    Write-Host "Skipped existing report/date checks: $skippedCount"
  }

  $sftpArgs = @(
    "-P", $Port,
    "-i", $PrivateKey,
    "-o", "StrictHostKeyChecking=accept-new",
    "-b", $batchFile.FullName,
    "$Username@$HostName"
  )

  & $SftpExecutable @sftpArgs
  $exitCode = $LASTEXITCODE
  if ($exitCode -ne 0) {
    throw "SFTP download failed with exit code $exitCode"
  }

  $downloaded = Get-ChildItem -LiteralPath $LocalDir -Filter "*.xlsx" -File -ErrorAction SilentlyContinue
  Write-Host "Local XLSX files now available: $($downloaded.Count)"
} finally {
  Remove-Item -LiteralPath $batchFile.FullName -Force -ErrorAction SilentlyContinue
}
