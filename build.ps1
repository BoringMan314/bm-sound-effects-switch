param(
  [string]$PyWin10 = "py -3.13",
  [string]$PyWin7 = "py -3.8"
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

function Ensure-CleanTempDirs {
  foreach ($d in @("build", "dist")) {
    $p = Join-Path $PSScriptRoot $d
    if (-not (Test-Path -LiteralPath $p)) {
      New-Item -ItemType Directory -Path $p | Out-Null
    }
    Get-ChildItem -LiteralPath $p -Force -ErrorAction SilentlyContinue | ForEach-Object {
      Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}

function Remove-OldExe {
  param([string]$OutputName)
  $exeName = "$OutputName.exe"
  $exePath = Join-Path $PSScriptRoot $exeName
  try {
    taskkill /F /IM $exeName /T *> $null
  } catch {}
  if (Test-Path -LiteralPath $exePath) {
    Remove-Item -LiteralPath $exePath -Force -ErrorAction Stop
    if (Test-Path -LiteralPath $exePath) {
      throw "Output exe is still locked: $exePath"
    }
  }
}

function Get-PythonVersionText {
  param([string]$PythonCmd)
  $cmd = "$PythonCmd -c `"import platform; print(platform.python_version())`""
  $output = $null
  try {
    $output = Invoke-Expression $cmd 2>$null
  } catch {
    throw "Cannot execute Python command: [$PythonCmd]. Please install the requested Python version first."
  }
  if ($LASTEXITCODE -ne 0 -or $null -eq $output -or [string]::IsNullOrWhiteSpace(($output | Out-String))) {
    throw "Cannot execute Python command: [$PythonCmd]. Please install the requested Python version first."
  }
  return (($output | Out-String).Trim())
}

function Assert-Python38Plus {
  param([string]$PythonCmd, [string]$Label)
  $cmd = "$PythonCmd -c `"import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}.{sys.version_info[2]}')`""
  $output = $null
  try {
    $output = Invoke-Expression $cmd 2>$null
  } catch {
    throw "[$Label] Cannot execute Python: [$PythonCmd]"
  }
  if ($LASTEXITCODE -ne 0 -or $null -eq $output -or [string]::IsNullOrWhiteSpace(($output | Out-String))) {
    throw "[$Label] Cannot execute Python: [$PythonCmd]"
  }
  $ver = (($output | Out-String).Trim())
  $parts = $ver.Split(".")
  if ($parts.Count -lt 2) {
    throw "[$Label] Cannot parse version: $ver"
  }
  $major = [int]$parts[0]
  $minor = [int]$parts[1]
  if ($major -lt 3 -or ($major -eq 3 -and $minor -lt 8)) {
    throw "[$Label] Requires Python 3.8+ (current $ver, cmd: [$PythonCmd])"
  }
}

function Assert-Python310Plus {
  param([string]$PythonCmd, [string]$Label)
  $cmd = "$PythonCmd -c `"import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}.{sys.version_info[2]}')`""
  $output = $null
  try {
    $output = Invoke-Expression $cmd 2>$null
  } catch {
    throw "[$Label] Cannot execute Python: [$PythonCmd]"
  }
  if ($LASTEXITCODE -ne 0 -or $null -eq $output -or [string]::IsNullOrWhiteSpace(($output | Out-String))) {
    throw "[$Label] Cannot execute Python: [$PythonCmd]"
  }
  $ver = (($output | Out-String).Trim())
  $parts = $ver.Split(".")
  if ($parts.Count -lt 2) {
    throw "[$Label] Cannot parse version: $ver"
  }
  $major = [int]$parts[0]
  $minor = [int]$parts[1]
  if ($major -lt 3 -or ($major -eq 3 -and $minor -lt 10)) {
    throw "[$Label] Requires Python 3.10+ (current $ver, cmd: [$PythonCmd])"
  }
}

function Invoke-BuildInVenv {
  param(
    [string]$PythonCmd,
    [string]$VenvDir,
    [string]$RequirementsFile,
    [string]$BuildMode,
    [string]$OutputName
  )

  $pythonVer = Get-PythonVersionText -PythonCmd $PythonCmd
  Write-Host "==> Build $OutputName with [$PythonCmd] (Python $pythonVer)"
  Remove-OldExe -OutputName $OutputName
  Ensure-CleanTempDirs

  Invoke-Expression "$PythonCmd -m venv --clear `"$VenvDir`""
  & "$VenvDir\Scripts\python.exe" -m pip install --upgrade pip
  & "$VenvDir\Scripts\python.exe" -m pip install -r $RequirementsFile
  & "$VenvDir\Scripts\python.exe" (Join-Path $PSScriptRoot "build.py") $BuildMode
  if ($LASTEXITCODE -ne 0) {
    throw "build.py $BuildMode failed for $OutputName"
  }
  $rootExe = Join-Path $PSScriptRoot "$OutputName.exe"
  if (-not (Test-Path -LiteralPath $rootExe)) {
    throw "Missing built exe: $rootExe"
  }
  Ensure-CleanTempDirs
}

Assert-Python310Plus -PythonCmd $PyWin10 -Label "Win10/11 build"
Assert-Python38Plus -PythonCmd $PyWin7 -Label "Win7 build chain"

Invoke-BuildInVenv `
  -PythonCmd $PyWin10 `
  -VenvDir ".venv-build-win10" `
  -RequirementsFile "requirements-win10.txt" `
  -BuildMode "win10" `
  -OutputName "bm-sound-effects-switch"

Invoke-BuildInVenv `
  -PythonCmd $PyWin7 `
  -VenvDir ".venv-build-win7" `
  -RequirementsFile "requirements-win7.txt" `
  -BuildMode "win7" `
  -OutputName "bm-sound-effects-switch_win7"

Write-Host ""
Write-Host "Build done. EXE paths:"
Write-Host "  .\bm-sound-effects-switch.exe (Win10/11 chain)"
Write-Host "  .\bm-sound-effects-switch_win7.exe (Win7 deps chain, see requirements-win7.txt)"
