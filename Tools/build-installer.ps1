[CmdletBinding()]
param(
    [string]$IsccPath,
    [string]$InstallerScript = (Join-Path $PSScriptRoot '..\Installer\OofManager.iss'),
    [string]$PowerAppsCliMsiPath,
    [switch]$SkipPowerAppsCliMsi,
    [switch]$SkipPublish,
    [string]$Configuration = 'Release',
    [string]$Project = (Join-Path $PSScriptRoot '..\OofManager.Wpf.csproj'),
    [string]$PublishDir = (Join-Path $PSScriptRoot '..\publish')
)

$ErrorActionPreference = 'Stop'

$powerAppsCliMsiUrl = 'https://download.microsoft.com/download/D/B/E/DBE69906-B4DA-471C-8960-092AB955C681/powerapps-cli-1.0.msi'
$powerAppsCliMsiSha256 = 'AF4C7058AB52E433ED0E18CB5D5E17256187A1B08ADD319F94C06C3EAC5D9190'

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

function Copy-PowerAppsCliMsi {
    param(
        [string]$DestinationPath,
        [string]$ExplicitPath
    )

    $destinationDir = Split-Path -Parent $DestinationPath
    if (-not (Test-Path -LiteralPath $destinationDir -PathType Container)) {
        New-Item -ItemType Directory -Path $destinationDir | Out-Null
    }

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        if (-not (Test-Path -LiteralPath $ExplicitPath -PathType Leaf)) {
            throw "PowerApps CLI MSI was not found at the explicit path: $ExplicitPath"
        }

        Copy-Item -LiteralPath $ExplicitPath -Destination $DestinationPath -Force
        return
    }

    $needsDownload = $true
    if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) {
        $hash = (Get-FileHash -LiteralPath $DestinationPath -Algorithm SHA256).Hash
        $needsDownload = -not [string]::Equals($hash, $powerAppsCliMsiSha256, [StringComparison]::OrdinalIgnoreCase)
    }

    if ($needsDownload) {
        Write-Host "Downloading Microsoft PowerApps CLI prerequisite..."
        Invoke-WebRequest -Uri $powerAppsCliMsiUrl -OutFile $DestinationPath -UseBasicParsing
    }

    $actualHash = (Get-FileHash -LiteralPath $DestinationPath -Algorithm SHA256).Hash
    if (-not [string]::Equals($actualHash, $powerAppsCliMsiSha256, [StringComparison]::OrdinalIgnoreCase)) {
        throw "PowerApps CLI MSI hash mismatch. Expected $powerAppsCliMsiSha256 but got $actualHash at $DestinationPath"
    }
}

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$installerScriptPath = (Resolve-Path -LiteralPath $InstallerScript).Path
$compilerPath = Resolve-IsccPath -ExplicitPath $IsccPath
$outputPath = Join-Path $repoRoot 'Installer\Output\OofManagerSetup.exe'
$powerAppsCliMsiOutputPath = Join-Path $repoRoot 'Installer\Prerequisites\powerapps-cli-1.0.msi'

Push-Location $repoRoot
try {
    if (-not $SkipPublish) {
        $projectPath = (Resolve-Path -LiteralPath $Project).Path
        Write-Host "Publishing $projectPath ($Configuration) -> $PublishDir"
        & dotnet publish $projectPath -c $Configuration -o $PublishDir -v minimal
        if ($LASTEXITCODE -ne 0) {
            throw "dotnet publish failed with exit code $LASTEXITCODE"
        }
    }

    if (-not $SkipPowerAppsCliMsi) {
        Copy-PowerAppsCliMsi -DestinationPath $powerAppsCliMsiOutputPath -ExplicitPath $PowerAppsCliMsiPath
        Write-Host "PowerApps CLI prerequisite: $powerAppsCliMsiOutputPath"
    }

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