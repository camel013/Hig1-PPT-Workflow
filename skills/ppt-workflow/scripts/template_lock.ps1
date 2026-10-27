param(
    [ValidateSet("Write", "Verify")]
    [string]$Mode = "Verify",
    [string]$Workspace = ""
)

$ErrorActionPreference = "Stop"
if (-not $Workspace) {
    $Workspace = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..\..\..")).Path
}
$root = (Resolve-Path -LiteralPath $Workspace).Path
$manifestPath = Join-Path $PSScriptRoot "..\references\template-lock.json"
$checksumPath = Join-Path $PSScriptRoot "..\references\template-lock.sha256"
$excludedRoots = @(
    (Join-Path $root ".git"),
    (Join-Path $root ".local")
)

function Test-ExcludedPath {
    param([string]$Path)
    foreach ($excluded in $excludedRoots) {
        if ($Path.StartsWith($excluded, [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function Get-LockedFile {
    param([System.IO.FileInfo]$File)
    $relative = [System.IO.Path]::GetRelativePath($root, $File.FullName).Replace("\", "/")
    $group = if ($relative.StartsWith("projects/")) {
        "project"
    }
    elseif ($relative.StartsWith("skills/")) {
        "skill"
    }
    elseif ($relative.StartsWith("design-system/")) {
        "design_system"
    }
    else {
        "workspace"
    }
    [ordered]@{
        group = $group
        path = $relative
        sha256 = (Get-FileHash -LiteralPath $File.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    }
}

if ($Mode -eq "Write") {
    $manifestRelative = [System.IO.Path]::GetRelativePath($root, [System.IO.Path]::GetFullPath($manifestPath)).Replace("\", "/")
    $checksumRelative = [System.IO.Path]::GetRelativePath($root, [System.IO.Path]::GetFullPath($checksumPath)).Replace("\", "/")
    $tracked = @(
        git -c "safe.directory=$root" -C $root -c core.quotepath=false ls-files |
            ForEach-Object { $_.Trim('"') }
    )
    if ($LASTEXITCODE -ne 0 -or $tracked.Count -eq 0) {
        throw "Unable to enumerate tracked files from Git."
    }
    $files = @(
        $tracked |
            Where-Object { $_ -ne $manifestRelative -and $_ -ne $checksumRelative } |
            Sort-Object |
            ForEach-Object { Get-LockedFile (Get-Item -LiteralPath (Join-Path $root $_)) }
    )
    $manifest = [ordered]@{
        schema_version = 2
        status = "active"
        frozen_at = (Get-Date).ToString("o")
        root = "."
        files = $files
    }
    New-Item -ItemType Directory -Path (Split-Path -Parent $manifestPath) -Force | Out-Null
    $manifest | ConvertTo-Json -Depth 6 |
        Set-Content -LiteralPath $manifestPath -Encoding UTF8
    $lines = foreach ($item in $files) {
        "$($item.sha256)  $($item.path)"
    }
    Set-Content -LiteralPath $checksumPath -Value $lines -Encoding UTF8
    [ordered]@{
        status = "WRITTEN"
        locked_files = $files.Count
        manifest = $manifestPath
        checksums = $checksumPath
    } | ConvertTo-Json
    exit 0
}

if (-not (Test-Path -LiteralPath $manifestPath)) {
    throw "Template lock manifest is missing: $manifestPath"
}

$stored = Get-Content -LiteralPath $manifestPath -Encoding UTF8 -Raw |
    ConvertFrom-Json
$errors = @()
foreach ($item in $stored.files) {
    $path = Join-Path $root $item.path
    if (-not (Test-Path -LiteralPath $path)) {
        $errors += [ordered]@{ type = "missing"; path = $item.path }
        continue
    }
    $actual = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actual -ne $item.sha256) {
        $errors += [ordered]@{
            type = "changed"
            path = $item.path
            expected = $item.sha256
            actual = $actual
        }
    }
}

[ordered]@{
    status = if ($errors.Count -eq 0) { "PASS" } else { "FAIL" }
    locked_files = @($stored.files).Count
    errors = $errors
} | ConvertTo-Json -Depth 5
if ($errors.Count -gt 0) {
    exit 1
}
