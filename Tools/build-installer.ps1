[CmdletBinding()]
param(
    [string]$IsccPath,
    [string]$InstallerScript = (Join-Path $PSScriptRoot '..\Installer\OofManager.iss')
)

$ErrorActionPreference = 'Stop'

function Resolve-IsccPath {
    param([string]$ExplicitPath)

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        if (Test-Path -LiteralPath $ExplicitPath -PathType Leaf) {
            return (Resolve-Path -LiteralPath $ExplicitPath).Path
        }

        throw "ISCC.exe was not found at the explicit path: $ExplicitPath"
    }

    $candidates = [System.Collections.Generic.List[string]]::new()

    if (-not [string]::IsNullOrWhiteSpace($env:ISCC_PATH)) {
        $candidates.Add($env:ISCC_PATH)
    }

    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        $candidates.Add((Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'))
        $candidates.Add((Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 5\ISCC.exe'))
    }

    foreach ($command in @(Get-Command ISCC.exe -ErrorAction SilentlyContinue)) {
        if (-not [string]::IsNullOrWhiteSpace($command.Source)) {
            $candidates.Add($command.Source)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
        $candidates.Add((Join-Path $env:ProgramFiles 'Inno Setup 6\ISCC.exe'))
        $candidates.Add((Join-Path $env:ProgramFiles 'Inno Setup 5\ISCC.exe'))
    }

    $programFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    if (-not [string]::IsNullOrWhiteSpace($programFilesX86)) {
        $candidates.Add((Join-Path $programFilesX86 'Inno Setup 6\ISCC.exe'))
        $candidates.Add((Join-Path $programFilesX86 'Inno Setup 5\ISCC.exe'))
    }

    $checked = [System.Collections.Generic.List[string]]::new()
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate) -or $checked.Contains($candidate)) {
            continue
        }

        $checked.Add($candidate)
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    $checkedText = ($checked | ForEach-Object { "  - $_" }) -join [Environment]::NewLine
    throw "ISCC.exe was not found. Install Inno Setup 6 with: winget install --id JRSoftware.InnoSetup -e`nChecked:`n$checkedText"
}

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$installerScriptPath = (Resolve-Path -LiteralPath $InstallerScript).Path
$compilerPath = Resolve-IsccPath -ExplicitPath $IsccPath
$outputPath = Join-Path $repoRoot 'Installer\Output\OofManagerSetup.exe'

Push-Location $repoRoot
try {
    Write-Host "Using Inno Setup compiler: $compilerPath"
    Write-Host "Building installer script: $installerScriptPath"

    & $compilerPath $installerScriptPath
    if ($LASTEXITCODE -ne 0) {
        throw "ISCC.exe failed with exit code $LASTEXITCODE"
    }

    if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
        throw "Installer output was not found: $outputPath"
    }

    $item = Get-Item -LiteralPath $outputPath
    $hash = Get-FileHash -LiteralPath $outputPath -Algorithm SHA256

    Write-Host "Installer: $($item.FullName)"
    Write-Host "Size: $($item.Length) bytes"
    Write-Host "SHA256: $($hash.Hash)"
}
finally {
    Pop-Location
}