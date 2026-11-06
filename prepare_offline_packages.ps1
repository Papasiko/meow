$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$Root = $PSScriptRoot
Set-Location -LiteralPath $Root
$Wheelhouse = Join-Path $Root 'wheelhouse'
$Requirements = Join-Path $Root 'requirements-offline-win312.txt'
$Builder = Join-Path $Root 'build_offline_zip.py'
$OutputZip = Join-Path (Split-Path -Parent $Root) 'OrgEvaluation_Windows10_Python312_OfflineReady.zip'

function Write-Info([string]$Message) { Write-Host $Message -ForegroundColor Cyan }
function Write-Fail([string]$Message) { Write-Host "ERROR: $Message" -ForegroundColor Red }

try {
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host 'Offline package preparation for Windows x64 / Python 3.12' -ForegroundColor Cyan
    Write-Host 'Run this on a Windows computer with Internet access.' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor Cyan

    if (-not (Test-Path -LiteralPath $Requirements)) { throw 'requirements-offline-win312.txt was not found.' }
    if (-not (Test-Path -LiteralPath $Builder)) { throw 'build_offline_zip.py was not found.' }

    $PythonExe = $null
    $PythonPrefix = @()
    $PyLauncher = Get-Command 'py.exe' -ErrorAction SilentlyContinue
    if ($PyLauncher) {
        & $PyLauncher.Source -3 -m pip --version *> $null
        if ($LASTEXITCODE -eq 0) { $PythonExe = $PyLauncher.Source; $PythonPrefix = @('-3') }
    }
    if (-not $PythonExe) {
        $PythonCommand = Get-Command 'python.exe' -ErrorAction SilentlyContinue
        if (-not $PythonCommand) { $PythonCommand = Get-Command 'python' -ErrorAction SilentlyContinue }
        if ($PythonCommand) {
            & $PythonCommand.Source -m pip --version *> $null
            if ($LASTEXITCODE -eq 0) { $PythonExe = $PythonCommand.Source; $PythonPrefix = @() }
        }
    }
    if (-not $PythonExe) { throw 'Python with pip was not found. Install Python with pip on this Internet-connected computer, then retry.' }

    Write-Info "Using Python command: $PythonExe $($PythonPrefix -join ' ')"
    $PipVersionArgs = @($PythonPrefix) + @('-m','pip','--version')
    & $PythonExe @PipVersionArgs
    if ($LASTEXITCODE -ne 0) { throw 'Could not run pip.' }

    if (-not (Test-Path -LiteralPath $Wheelhouse)) { New-Item -ItemType Directory -Path $Wheelhouse | Out-Null }
    Get-ChildItem -LiteralPath $Wheelhouse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -eq '.whl' -or $_.Name.EndsWith('.tar.gz') } | Remove-Item -Force

    $BaseArgs = @('-m','pip','download','--timeout','120','--retries','2','--only-binary=:all:','--platform','win_amd64','--python-version','3.12','--implementation','cp','--abi','cp312','--abi','abi3','--abi','none','--dest',$Wheelhouse,'-r',$Requirements)
    Write-Host ''
    Write-Info 'Downloading Windows x64 / CPython 3.12 wheels from the default Python package index...'
    $DownloadArgs = @($PythonPrefix) + $BaseArgs
    & $PythonExe @DownloadArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Host ''
        Write-Host 'Default source failed. Retrying with the Aliyun mirror...' -ForegroundColor Yellow
        Get-ChildItem -LiteralPath $Wheelhouse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -eq '.whl' -or $_.Name.EndsWith('.tar.gz') } | Remove-Item -Force
        $MirrorArgs = @('-m','pip','download','--timeout','120','--retries','3','--trusted-host','mirrors.aliyun.com','-i','https://mirrors.aliyun.com/pypi/simple/','--only-binary=:all:','--platform','win_amd64','--python-version','3.12','--implementation','cp','--abi','cp312','--abi','abi3','--abi','none','--dest',$Wheelhouse,'-r',$Requirements)
        $DownloadArgs = @($PythonPrefix) + $MirrorArgs
        & $PythonExe @DownloadArgs
        if ($LASTEXITCODE -ne 0) { throw 'Package download failed from both sources. Check Internet access and proxy/firewall settings.' }
    }

    $WheelCount = @(Get-ChildItem -LiteralPath $Wheelhouse -Filter '*.whl' -File).Count
    Write-Host "Wheel files downloaded: $WheelCount"
    if ($WheelCount -lt 8) { throw "The wheelhouse appears incomplete ($WheelCount wheels found; at least 8 expected)." }

    Write-Host ''
    Write-Info 'Building and checking the offline-ready ZIP...'
    $BuilderArgs = @($PythonPrefix) + @($Builder)
    & $PythonExe @BuilderArgs
    if ($LASTEXITCODE -ne 0) { throw 'build_offline_zip.py failed.' }
    if (-not (Test-Path -LiteralPath $OutputZip)) { throw 'The final ZIP was not created.' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $Archive = [System.IO.Compression.ZipFile]::OpenRead($OutputZip)
    try {
        $WheelEntries = @($Archive.Entries | Where-Object { $_.FullName -match '/wheelhouse/[^/]+\.whl$' }).Count
        if ($WheelEntries -lt 8) { throw "ZIP verification failed: only $WheelEntries wheel files found in the archive." }
    } finally { $Archive.Dispose() }

    Write-Host ''
    Write-Host 'SUCCESS: Offline-ready ZIP created and checked.' -ForegroundColor Green
    Write-Host "File: $OutputZip" -ForegroundColor Green
    Write-Host 'Copy this ZIP to a USB drive and take it to the offline server.' -ForegroundColor Green
    exit 0
}
catch {
    Write-Fail $_.Exception.Message
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    exit 1
}
finally {
    Write-Host ''
}
