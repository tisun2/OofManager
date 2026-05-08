# Generates winget manifests for a given OofManager release and (optionally)
# pushes them to a forked microsoft/winget-pkgs repo as a PR.
#
# Examples:
#   # Just regenerate manifests for the current installer:
#   .\Tools\winget-publish.ps1 -Version 1.0.9
#
#   # Generate + open PR:
#   .\Tools\winget-publish.ps1 -Version 1.0.9 -CreatePR
#
# Required for -CreatePR:
#   * Local clone of the fork at $WingetForkPath (defaults to $HOME\winget-pkgs)
#   * The fork has `upstream` pointing at microsoft/winget-pkgs
#   * git is configured to push to your fork over HTTPS or SSH
#
# Notes:
#   * The installer must already be built and uploaded to the matching GitHub
#     release URL (https://github.com/<Owner>/<Repo>/releases/download/v<Version>/<InstallerName>)
#     before -CreatePR runs; the winget bot resolves the URL during PR validation.
#   * The script does NOT call `gh pr create` because microsoft/winget-pkgs is
#     SAML-protected and rejects unauthorized OAuth tokens at the GraphQL layer.
#     Instead it opens the GitHub compare URL in your browser; click "Create
#     pull request" there. (Existing browser auth handles SAML transparently.)
#   * The upstream fetch is shallow (`--depth=1 master --no-tags`) so the
#     ~500 MB winget-pkgs history is not downloaded on every release.
#   * Re-running with the same version overwrites the local manifests but
#     does NOT touch a previously merged PR — bump the version first.

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [string]$RepoRoot,
    [string]$InstallerPath,
    [string]$InstallerName = 'OofManagerSetup.exe',
    [string]$PackageIdentifier = 'tisun2.OofManager',
    [string]$Owner = 'tisun2',
    [string]$Repo = 'OofManager',
    [string]$ReleaseDate = (Get-Date -Format 'yyyy-MM-dd'),
    [string]$WingetForkPath = (Join-Path $env:USERPROFILE 'winget-pkgs'),

    [switch]$CreatePR
)

$ErrorActionPreference = 'Stop'

# Resolve defaults inside the body. Doing it in [param()] defaults yields an
# empty $PSScriptRoot under some hosts (advanced functions, ScriptBlock-style
# invocation), which would silently break the script with "Path is empty".
if (-not $RepoRoot) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
}
if (-not $InstallerPath) {
    $InstallerPath = Join-Path $RepoRoot 'Installer\Output\OofManagerSetup.exe'
}

if (-not (Test-Path $InstallerPath)) {
    throw "Installer not found at $InstallerPath. Build the installer first (publish + ISCC), then re-run."
}

# --- Compute installer hash ---
Write-Host "Hashing $InstallerPath ..."
$sha256 = (Get-FileHash $InstallerPath -Algorithm SHA256).Hash
Write-Host "  SHA256: $sha256"

# --- Write manifests ---
$manifestDir = Join-Path $RepoRoot "winget\$Version"
New-Item -ItemType Directory -Path $manifestDir -Force | Out-Null

# Best-effort split of "tisun2.OofManager" into Publisher + PackageName parts
# for the directory tree microsoft/winget-pkgs expects:
#   manifests/<lowercase first letter of Publisher>/<Publisher>/<PackageName>/<Version>/
$idParts = $PackageIdentifier -split '\.', 2
if ($idParts.Length -lt 2) {
    throw "PackageIdentifier '$PackageIdentifier' must be at least Publisher.PackageName."
}
$publisherSegment = $idParts[0]
$packageSegment   = $idParts[1]
$firstLetter      = $publisherSegment.Substring(0, 1).ToLowerInvariant()

# version.yaml
@"
# yaml-language-server: `$schema=https://aka.ms/winget-manifest.version.1.10.0.schema.json

PackageIdentifier: $PackageIdentifier
PackageVersion: $Version
DefaultLocale: en-US
ManifestType: version
ManifestVersion: 1.10.0
"@ | Set-Content -Encoding utf8 (Join-Path $manifestDir "$PackageIdentifier.yaml")

# installer.yaml
@"
# yaml-language-server: `$schema=https://aka.ms/winget-manifest.installer.1.10.0.schema.json

PackageIdentifier: $PackageIdentifier
PackageVersion: $Version
MinimumOSVersion: 10.0.0.0
InstallerType: inno
Scope: user
InstallModes:
  - interactive
  - silent
  - silentWithProgress
InstallerSwitches:
  Silent: /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
  SilentWithProgress: /SILENT /SUPPRESSMSGBOXES /NORESTART
UpgradeBehavior: install
ReleaseDate: $ReleaseDate
Installers:
  - Architecture: x64
    InstallerUrl: https://github.com/$Owner/$Repo/releases/download/v$Version/$InstallerName
    InstallerSha256: $sha256
ManifestType: installer
ManifestVersion: 1.10.0
"@ | Set-Content -Encoding utf8 (Join-Path $manifestDir "$PackageIdentifier.installer.yaml")

# locale.en-US.yaml
@"
# yaml-language-server: `$schema=https://aka.ms/winget-manifest.defaultLocale.1.10.0.schema.json

PackageIdentifier: $PackageIdentifier
PackageVersion: $Version
PackageLocale: en-US
Publisher: $publisherSegment
PublisherUrl: https://github.com/$Owner
PublisherSupportUrl: https://github.com/$Owner/$Repo/issues
PackageName: OOF Manager
PackageUrl: https://github.com/$Owner/$Repo
License: MIT
LicenseUrl: https://github.com/$Owner/$Repo/blob/main/LICENSE
Copyright: Copyright (c) $publisherSegment
ShortDescription: Manage Microsoft 365 / Exchange Online Out-of-Office (Automatic Reply) settings without opening Outlook.
Description: |-
  OOF Manager is a lightweight Windows desktop tool that manages your Microsoft 365 mailbox's
  Out-of-Office (Automatic Reply) state for you. Sign in once, configure your weekly work hours
  and reply text, and OOF Manager auto-toggles your OOF on/off at the schedule boundaries.
  It also pushes the next off-hours window to Exchange as a Scheduled OOF, so the server keeps
  switching auto-replies even when this app is closed.
Moniker: oofmanager
Tags:
  - exchange
  - microsoft-365
  - office-365
  - out-of-office
  - oof
  - outlook
  - autoreply
ReleaseNotesUrl: https://github.com/$Owner/$Repo/releases/tag/v$Version
ManifestType: defaultLocale
ManifestVersion: 1.10.0
"@ | Set-Content -Encoding utf8 (Join-Path $manifestDir "$PackageIdentifier.locale.en-US.yaml")

Write-Host "Wrote manifests to $manifestDir"

# --- Local validation ---
Write-Host "Running winget validate ..."
& winget validate --manifest $manifestDir
if ($LASTEXITCODE -ne 0) {
    throw "winget validate failed. Fix the manifests above before publishing."
}

if (-not $CreatePR) {
    Write-Host ""
    Write-Host "Done. Re-run with -CreatePR to push these manifests as a PR."
    return
}

# --- Push manifests to forked winget-pkgs and open PR ---

if (-not (Test-Path $WingetForkPath)) {
    throw @"
Fork of microsoft/winget-pkgs not found at $WingetForkPath.
Set up once with:
  gh repo clone <YourGitHubUser>/winget-pkgs $WingetForkPath
  cd $WingetForkPath
  git remote add upstream https://github.com/microsoft/winget-pkgs.git
"@
}

# Layout under microsoft/winget-pkgs is:
#   manifests/<firstLetter>/<Publisher>/<PackageName>/<Version>/...
$destDir = Join-Path $WingetForkPath "manifests\$firstLetter\$publisherSegment\$packageSegment\$Version"
Write-Host "Copying manifests to $destDir"
New-Item -ItemType Directory -Path $destDir -Force | Out-Null
Copy-Item (Join-Path $manifestDir '*') $destDir -Force

Push-Location $WingetForkPath
try {
    # Fetch only the tip of upstream/master, with no tags and shallow depth.
    # The default `git fetch upstream` pulls every branch in microsoft/winget-pkgs
    # including hundreds of stale bot branches (~500 MB / 2.5M deltas) and
    # takes minutes. The winget bot rebases our PR onto current master at
    # validation time, so the only commit we actually need locally is the
    # current master tip.
    Write-Host "Fetching upstream/master tip (shallow) ..."
    & git fetch upstream master --depth=1 --no-tags
    if ($LASTEXITCODE -ne 0) {
        throw "git fetch upstream master failed. Aborting."
    }
    $branchName = "$PackageIdentifier-$Version"
    & git checkout -B $branchName FETCH_HEAD | Out-Null

    # Re-copy AFTER the branch checkout because checkout -B reset the worktree.
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    Copy-Item (Join-Path $manifestDir '*') $destDir -Force

    & git add (Resolve-Path "manifests\$firstLetter\$publisherSegment\$packageSegment\$Version")
    & git commit -m "New version: $PackageIdentifier version $Version"
    if ($LASTEXITCODE -ne 0) {
        throw "git commit failed. Aborting."
    }
    & git push -u origin $branchName
    if ($LASTEXITCODE -ne 0) {
        throw "git push failed. Aborting."
    }

    # Use --web instead of `gh pr create` so we don't hit the GraphQL endpoint,
    # which microsoft/winget-pkgs blocks with "Resource protected by organization
    # SAML enforcement" for tokens that haven't been SAML-authorized for the
    # microsoftopensource enterprise. The compare-url path doesn't need the
    # API at all and works for any auth state the user already has in the
    # browser.
    Write-Host "Opening PR creation page in browser ..."
    Start-Process "https://github.com/microsoft/winget-pkgs/compare/master...$($Owner):winget-pkgs:$branchName?expand=1&title=$([uri]::EscapeDataString("New version: $PackageIdentifier version $Version"))"
}
finally {
    Pop-Location
}
