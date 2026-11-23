Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Section {
  param([string]$Message)
  Write-Host "`n=== $Message ===" -ForegroundColor Cyan
}

function Write-Step {
  param([string]$Message)
  Write-Host "[step] $Message" -ForegroundColor DarkCyan
}

function Write-Ok {
  param([string]$Message)
  Write-Host "[ok] $Message" -ForegroundColor Green
}

function Write-Warn {
  param([string]$Message)
  Write-Host "[warn] $Message" -ForegroundColor Yellow
}

function Write-Info {
  param([string]$Message)
  Write-Host "[info] $Message" -ForegroundColor Gray
}

function Assert-Command {
  param([string]$CommandName)
  $cmd = Get-Command $CommandName -ErrorAction SilentlyContinue
  if (-not $cmd) {
    throw "Required command not found in PATH: $CommandName"
  }
  return $true
}

function Test-CommandExists {
  param([string]$CommandName)
  return [bool](Get-Command $CommandName -ErrorAction SilentlyContinue)
}

function Refresh-ProcessPath {
  $machinePath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
  $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
  $env:Path = "$machinePath;$userPath"
}

function Invoke-External {
  param(
    [Parameter(Mandatory = $true)][string]$FilePath,
    [string[]]$ArgumentList = @(),
    [string]$WorkingDirectory = '',
    # Some steps (npm install, Electron's postinstall binary download, etc.)
    # fetch large files from GitHub-hosted CDNs, which have been observed to
    # intermittently drop the TLS connection mid-download on some corporate
    # networks. Retrying the same command a couple of times clears this
    # without any code change on the network side.
    [int]$MaxAttempts = 1,
    [int]$RetryDelaySeconds = 5
  )

  for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
    try {
      if ($WorkingDirectory) {
        Push-Location $WorkingDirectory
      }
      & $FilePath @ArgumentList
      if ($LASTEXITCODE -ne 0) {
        throw "Command failed ($LASTEXITCODE): $FilePath $($ArgumentList -join ' ')"
      }
      return
    }
    catch {
      if ($attempt -ge $MaxAttempts) { throw }
      Write-Warn "Attempt $attempt of $MaxAttempts failed ($($_.Exception.Message)) - retrying in $RetryDelaySeconds s..."
      Start-Sleep -Seconds $RetryDelaySeconds
    }
    finally {
      if ($WorkingDirectory) {
        Pop-Location
      }
    }
  }
}

function Ensure-Directory {
  param([Parameter(Mandatory = $true)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
  }
}

function Repair-ElectronBinary {
  param(
    [Parameter(Mandatory = $true)][string]$AppDesktopDir,
    [int]$MaxAttempts = 3,
    [int]$RetryDelaySeconds = 5
  )

  $electronDir = Join-Path $AppDesktopDir 'node_modules\electron'
  $exePath = Join-Path $electronDir 'dist\electron.exe'
  if (Test-Path -LiteralPath $exePath) {
    return
  }

  # Some endpoint-security products strip electron.exe out of the zip that
  # npm's own postinstall (@electron/get) downloads and extracts, while
  # leaving non-executable files (locales, etc.) untouched - and npm reports
  # success either way since it never re-checks the extraction result. A
  # plain Expand-Archive of the same zip has been confirmed to extract intact,
  # so fetch/extract it ourselves as a fallback instead of asking anyone to
  # change endpoint-security settings.
  Write-Warn 'Electron binary missing after npm install - extracting it manually instead...'

  $pkgJsonPath = Join-Path $electronDir 'package.json'
  if (-not (Test-Path -LiteralPath $pkgJsonPath)) {
    throw "Cannot repair Electron binary: $pkgJsonPath not found. Run npm install in $AppDesktopDir first."
  }
  $version = (Get-Content -LiteralPath $pkgJsonPath -Raw | ConvertFrom-Json).version
  $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'x64' }
  $zipName = "electron-v$version-win32-$arch.zip"
  $cacheDir = Join-Path $env:LOCALAPPDATA 'electron\Cache'
  Ensure-Directory -Path $cacheDir
  $zipPath = Join-Path $cacheDir $zipName

  if (-not (Test-Path -LiteralPath $zipPath)) {
    $url = "https://github.com/electron/electron/releases/download/v$version/$zipName"
    Write-Info "Downloading Electron binary directly: $url"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    # Same GitHub-hosted CDN that caused the original TLS-drop failure, so
    # this needs the same retry-with-backoff treatment as Invoke-External.
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
      try {
        Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $zipPath
        break
      }
      catch {
        Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
        if ($attempt -ge $MaxAttempts) { throw }
        Write-Warn "Download attempt $attempt of $MaxAttempts failed ($($_.Exception.Message)) - retrying in $RetryDelaySeconds s..."
        Start-Sleep -Seconds $RetryDelaySeconds
      }
    }
  } else {
    Write-Info "Reusing cached Electron zip: $zipPath"
  }

  $distDir = Join-Path $electronDir 'dist'
  Ensure-Directory -Path $distDir
  Expand-Archive -Path $zipPath -DestinationPath $distDir -Force

  # Checking electron.exe alone isn't enough - endpoint security could in
  # principle strip a different file (a required DLL/resource) and leave
  # electron.exe untouched. Compare the extracted file count against the
  # zip's own entry count so any partial strip is caught, not just this one.
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zipHandle = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
  try {
    $expectedFileCount = ($zipHandle.Entries | Where-Object { -not $_.FullName.EndsWith('/') }).Count
  } finally {
    $zipHandle.Dispose()
  }
  $actualFileCount = (Get-ChildItem -LiteralPath $distDir -Recurse -File -ErrorAction SilentlyContinue).Count

  if (-not (Test-Path -LiteralPath $exePath) -or $actualFileCount -lt $expectedFileCount) {
    throw "Electron binary extraction incomplete ($actualFileCount of $expectedFileCount files present, electron.exe " +
          "$(if (Test-Path -LiteralPath $exePath) { 'present' } else { 'missing' })). This looks like an antivirus/" +
          "endpoint-security block that needs an exclusion for this folder."
  }
  Write-Ok "Electron binary extracted successfully ($actualFileCount files)"
}

function Test-PlaywrightChromiumBinary {
  # Detection-only (unlike Repair-ElectronBinary, no confirmed safe repair
  # path yet for Playwright's own browser install) - just surfaces the same
  # class of silent endpoint-security strip early instead of it resurfacing
  # later as a confusing test failure.
  $browsersRoot = if ($env:PLAYWRIGHT_BROWSERS_PATH) { $env:PLAYWRIGHT_BROWSERS_PATH } else { Join-Path $env:LOCALAPPDATA 'ms-playwright' }
  if (-not (Test-Path -LiteralPath $browsersRoot)) {
    Write-Warn "Playwright browsers folder not found: $browsersRoot"
    return $false
  }
  $chromeExe = Get-ChildItem -LiteralPath $browsersRoot -Recurse -Filter 'chrome.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $chromeExe) {
    Write-Warn ("chrome.exe not found under $browsersRoot after install-browsers - this looks like the same " +
                'endpoint-security stripping seen with the Electron binary. Report parity tests will fail to launch a browser.')
    return $false
  }
  Write-Ok "Playwright chromium binary found: $($chromeExe.FullName)"
  return $true
}

function Copy-TemplateIfMissing {
  param(
    [Parameter(Mandatory = $true)][string]$DestinationPath,
    [Parameter(Mandatory = $true)][string]$TemplatePath
  )

  if (Test-Path -LiteralPath $DestinationPath) {
    return $false
  }

  if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "Template file not found: $TemplatePath"
  }

  Copy-Item -LiteralPath $TemplatePath -Destination $DestinationPath -Force
  return $true
}

function Read-PlainTextSecret {
  param([Parameter(Mandatory = $true)][string]$Prompt)

  $secure = Read-Host -AsSecureString -Prompt $Prompt
  $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
  try {
    return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
  }
  finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
  }
}

function Ensure-ToolWithWinget {
  param(
    [Parameter(Mandatory = $true)][string]$ToolName,
    [Parameter(Mandatory = $true)][string]$WingetId,
    [switch]$VerifyOnly
  )

  if (Test-CommandExists $ToolName) {
    Write-Ok "$ToolName is available"
    return
  }

  if ($VerifyOnly) {
    throw "$ToolName is missing"
  }

  if (-not (Test-CommandExists 'winget')) {
    throw "winget is required to install $ToolName automatically. Install $ToolName manually and rerun."
  }

  Write-Step "Installing $ToolName via winget ($WingetId)..."
  Invoke-External -FilePath 'winget' -ArgumentList @(
    'install', '--id', $WingetId, '-e',
    '--accept-source-agreements', '--accept-package-agreements'
  )

  Refresh-ProcessPath
  if (-not (Test-CommandExists $ToolName)) {
    throw "$ToolName still not found after installation. Close this window, open a new PowerShell window, and rerun installer."
  }

  Write-Ok "$ToolName installed"
}

function Test-Yes {
  param([string]$Value)
  if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
  $v = $Value.Trim().ToLowerInvariant()
  return ($v -eq 'y' -or $v -eq 'yes' -or $v -eq 'true' -or $v -eq '1')
}
