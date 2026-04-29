<#
.SYNOPSIS
Build and optionally publish a Windows release for RK3562 MCU UART Validation Tool.

.DESCRIPTION
By default, this script only builds local release artifacts.
Use -Push to commit source changes, push main, and create/push the release tag.
Use -CreateRelease together with -Push to publish a GitHub release.

.PARAMETER Version
Release version in the form 1.2.0 or v1.2.0.

.PARAMETER Notes
Optional GitHub release notes. If omitted, default notes are generated.

.PARAMETER RunTests
Runs python -m py_compile before packaging.

.PARAMETER Push
Stages allowed files, creates a release commit, pushes main, and creates/pushes the release tag.

.PARAMETER CreateRelease
Creates the GitHub release for the requested tag. Requires -Push.

.PARAMETER WhatIf
Prints the workflow steps without modifying files or publishing anything.

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File .\build-release.ps1 -Version 1.1.2
Build local artifacts only.

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File .\build-release.ps1 -Version 1.1.2 -RunTests
Build local artifacts after a syntax check.

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File .\build-release.ps1 -Version 1.1.2 -RunTests -Push
Build, commit, push main, and push tag v1.1.2.

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File .\build-release.ps1 -Version 1.1.2 -RunTests -Push -CreateRelease
Build, push, and publish the GitHub release.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [string]$Notes = "",
    [switch]$RunTests,
    [switch]$Push,
    [switch]$CreateRelease,
    [switch]$WhatIf,
    [string]$CommitMessage = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

$pythonFile = Join-Path $repoRoot "rk3562_uart_tester.py"
$pngFile = Join-Path $repoRoot "a.png"
$iconFile = Join-Path $repoRoot "a.ico"
$releaseDir = Join-Path $repoRoot "release"
$distDir = Join-Path $repoRoot "dist"
$buildDir = Join-Path $repoRoot "build"
$mainBranch = "main"
$appName = "rk3562_mcu_uart_validation_tool"
$appTitle = "RK3562 MCU UART Validation Tool"
$specFile = Join-Path $repoRoot "$appName.spec"

if ($Version -notmatch '^v?\d+\.\d+\.\d+$') {
    throw "Version must look like 1.2.0 or v1.2.0"
}

$normalizedVersion = $Version.TrimStart('v')
$tag = "v$normalizedVersion"
$releaseExeName = "${appName}_${tag}.exe"
$releaseZipName = "${appName}_${tag}.zip"
$distExe = Join-Path $distDir "$appName.exe"
$releaseExe = Join-Path $releaseDir $releaseExeName
$releaseZip = Join-Path $releaseDir $releaseZipName
$defaultNotes = @"
## Summary
- Release $tag for $appTitle

## Assets
- $releaseExeName
- $releaseZipName
"@
if ([string]::IsNullOrWhiteSpace($Notes)) {
    $Notes = $defaultNotes
}
$defaultCommitMessage = "Release version $normalizedVersion."

function Invoke-Step {
    param(
        [string]$Message,
        [scriptblock]$Action
    )

    Write-Host "==> $Message"
    if ($WhatIf) {
        return
    }
    & $Action
}

function Invoke-CheckedCommand {
    param(
        [string]$FilePath,
        [string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed: $FilePath $($Arguments -join ' ')"
    }
}

function Get-CommandPath {
    param([string]$Name)

    $command = Get-Command $Name -ErrorAction Stop
    return $command.Source
}

$pythonCmd = Get-CommandPath "python"
$pyinstallerCmd = Get-CommandPath "pyinstaller"
$gitCmd = Get-CommandPath "git"
$ghCmd = Get-CommandPath "gh"

$currentBranch = (& $gitCmd branch --show-current).Trim()
if ($LASTEXITCODE -ne 0) {
    throw "Unable to detect current git branch"
}
if ($currentBranch -ne $mainBranch) {
    throw "Release script must run on $mainBranch. Current branch: $currentBranch"
}

$originUrl = (& $gitCmd remote get-url origin).Trim()
if ($LASTEXITCODE -ne 0) {
    throw "Unable to read git remote origin"
}
if ([string]::IsNullOrWhiteSpace($originUrl)) {
    throw "Git remote origin is not configured"
}

$statusLines = & $gitCmd status --porcelain
if ($LASTEXITCODE -ne 0) {
    throw "Unable to read git status"
}
if ($statusLines) {
    $allowedDirty = @(
        " M rk3562_uart_tester.py",
        " M a.ico",
        " M a.png",
        " M bg.png",
        " M build-release.ps1",
        "?? bg.png",
        "?? build-release.ps1"
    )
    $unexpected = @($statusLines | Where-Object { $_ -notin $allowedDirty })
    if ($unexpected.Count -gt 0) {
        throw "Working tree has unrelated changes:`n$($unexpected -join "`n")"
    }
}

$source = Get-Content -Path $pythonFile -Raw -Encoding UTF8
$versionPattern = 'APP_VERSION = "[^"]+"'
if ($source -notmatch $versionPattern) {
    throw "APP_VERSION constant not found in rk3562_uart_tester.py"
}
$newSource = [regex]::Replace($source, $versionPattern, "APP_VERSION = `"$normalizedVersion`"", 1)

Invoke-Step "Update app version to $normalizedVersion" {
    [System.IO.File]::WriteAllText($pythonFile, $newSource, [System.Text.UTF8Encoding]::new($false))
}

if ($RunTests) {
    Invoke-Step "Run syntax check" {
        Invoke-CheckedCommand $pythonCmd @("-m", "py_compile", $pythonFile)
    }
}

Invoke-Step "Convert PNG icon to ICO" {
    $tmpPy = [System.IO.Path]::GetTempFileName() + ".py"
    @"
from PIL import Image
png = Image.open(r'$pngFile')
if png.mode != 'RGBA':
    png = png.convert('RGBA')
sizes = [16, 32, 48, 64, 128, 256]
imgs = []
for s in sizes:
    if s <= max(png.size):
        imgs.append(png.resize((s, s), Image.Resampling.LANCZOS))
if imgs:
    imgs[-1].save(r'$iconFile', format='ICO',
                 sizes=[(i.width, i.height) for i in imgs],
                 append_images=imgs[:-1])
"@ | Set-Content -Path $tmpPy -Encoding UTF8
    try {
        Invoke-CheckedCommand $pythonCmd @($tmpPy)
    } finally {
        Remove-Item -Path $tmpPy -Force -ErrorAction SilentlyContinue
    }
}

Invoke-Step "Build Windows executable" {
    Invoke-CheckedCommand $pyinstallerCmd @(
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--icon", $iconFile,
        "--name", $appName,
        "--add-data", "bg.png;.",
        $pythonFile
    )
}

Invoke-Step "Prepare release directory" {
    New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
    Remove-Item -Path $releaseExe, $releaseZip -Force -ErrorAction SilentlyContinue
    Copy-Item -Path $distExe -Destination $releaseExe -Force
    Compress-Archive -Path $releaseExe -DestinationPath $releaseZip -Force
}

Invoke-Step "Remove generated build artifacts" {
    Remove-Item -Path $buildDir, $distDir, $specFile -Recurse -Force -ErrorAction SilentlyContinue
}

if ($Push -or $CreateRelease) {
    Invoke-Step "Fetch remote tags" {
        Invoke-CheckedCommand $gitCmd @("fetch", "--tags", "origin")
    }

    $localTagExists = & $gitCmd rev-parse -q --verify "refs/tags/$tag" 2>$null
    if ($LASTEXITCODE -eq 0 -and $localTagExists) {
        throw "Tag $tag already exists locally"
    }

    $remoteTagExists = & $gitCmd ls-remote --tags origin $tag
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to query remote tags"
    }
    if (-not [string]::IsNullOrWhiteSpace($remoteTagExists)) {
        throw "Tag $tag already exists on origin"
    }
}

if ($Push) {
    Invoke-Step "Show files to commit" {
        & $gitCmd status --short -- $pythonFile $iconFile $pngFile (Join-Path $repoRoot "build-release.ps1")
        if ($LASTEXITCODE -ne 0) {
            throw "Unable to show pending source changes"
        }
    }

    if ([string]::IsNullOrWhiteSpace($CommitMessage)) {
        $CommitMessage = Read-Host "Commit message (Enter = '$defaultCommitMessage')"
        if ([string]::IsNullOrWhiteSpace($CommitMessage)) {
            $CommitMessage = $defaultCommitMessage
        }
    }

    Invoke-Step "Commit source changes" {
        Invoke-CheckedCommand $gitCmd @("add", "--", $pythonFile, $iconFile, $pngFile, (Join-Path $repoRoot "build-release.ps1"))
        Invoke-CheckedCommand $gitCmd @("commit", "-m", $CommitMessage)
    }

    Invoke-Step "Push branch $mainBranch" {
        Invoke-CheckedCommand $gitCmd @("push", "origin", $mainBranch)
    }

    Invoke-Step "Create release tag $tag" {
        Invoke-CheckedCommand $gitCmd @("tag", "-a", $tag, "-m", "Release $tag")
        Invoke-CheckedCommand $gitCmd @("push", "origin", $tag)
    }
}

if ($CreateRelease) {
    if (-not $Push) {
        throw "-CreateRelease requires -Push so the tag exists on GitHub"
    }

    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = "SilentlyContinue"
    & $ghCmd release view $tag 1>$null 2>$null
    $exit = $LASTEXITCODE
    $ErrorActionPreference = $prevEAP
    if ($exit -eq 0) {
        throw "GitHub release $tag already exists"
    }

    Invoke-Step "Create GitHub release $tag" {
        Invoke-CheckedCommand $ghCmd @(
            "release", "create", $tag,
            $releaseExe,
            $releaseZip,
            "--title", $tag,
            "--notes", $Notes
        )
    }
}

Write-Host "Release workflow completed for $tag"
if ($WhatIf) {
    Write-Host "WhatIf mode did not modify files or publish anything"
}
