param(
    [ValidateSet("Write", "Verify")]
    [string]$Mode = "Verify",
    [string]$Workspace = "C:\Users\camel\Documents\ChatGPT\Hig1",
    [string]$SkillRoot = "C:\Users\camel\.codex\skills\ppt-master-skill"
)

$ErrorActionPreference = "Stop"
$root = (Resolve-Path -LiteralPath $Workspace).Path
$manifestPath = Join-Path $root "TEMPLATE_LOCK.json"
$checksumPath = Join-Path $root "TEMPLATE_LOCK.sha256"

function Add-GroupFiles {
    param(
        [System.Collections.Generic.List[object]]$Target,
        [string]$Group,
        [object[]]$Files
    )
    foreach ($file in $Files) {
        if ($null -eq $file) {
            continue
        }
        $path = (Resolve-Path -LiteralPath $file.FullName).Path
        if ($Target.Path -contains $path) {
            continue
        }
        $Target.Add([pscustomobject]@{
            group = $Group
            absolute_path = $path
            sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        })
    }
}

function Collect-LockedFiles {
    $items = [System.Collections.Generic.List[object]]::new()
    Add-GroupFiles $items "workspace_instruction" @(Get-Item -LiteralPath (Join-Path $root "AGENTS.md"))
    Add-GroupFiles $items "workspace_tool" @(Get-Item -LiteralPath (Join-Path $root "template_lock.ps1"))
    $deliveryZip = Join-Path $root "减少无效会议_反馈闭环_重建包_20261026.zip"
    if (Test-Path -LiteralPath $deliveryZip) {
        Add-GroupFiles $items "workspace_delivery" @(Get-Item -LiteralPath $deliveryZip)
    }
    Add-GroupFiles $items "ppt_workflow" @(
        Get-ChildItem -LiteralPath (Join-Path $root "_ppt_workflow") -Recurse -File `
            -ErrorAction SilentlyContinue
    )
    Add-GroupFiles $items "duoyuhua_rebuild" @(
        Get-ChildItem -LiteralPath (Join-Path $root "_quality_baseline_duoyuhua_20261026") `
            -Recurse -File -ErrorAction SilentlyContinue
    )
    Add-GroupFiles $items "glass_sample" @(
        Get-ChildItem -LiteralPath (Join-Path $root "glass_layer_sample_20261026") `
            -Recurse -File -ErrorAction SilentlyContinue
    )

    $projectFinal = @(
        Get-ChildItem -LiteralPath (Join-Path $root "projects") -Recurse -File `
            -ErrorAction SilentlyContinue |
            Where-Object {
                $_.FullName -notmatch "\\(?:\.venv|vendor|backups|__pycache__|node_modules)\\" -and
                (
                    $_.Extension -in @(".pptx", ".pdf", ".zip", ".md") -or
                    $_.FullName -match "\\deliverables\\"
                )
            }
    )
    Add-GroupFiles $items "completed_projects" $projectFinal

    if (Test-Path -LiteralPath $SkillRoot) {
        $skillPath = (Resolve-Path -LiteralPath $SkillRoot).Path
        $skillFiles = @(
            Get-Item -LiteralPath (Join-Path $skillPath "SKILL.md")
            Get-ChildItem -LiteralPath (Join-Path $skillPath "agents") -File -ErrorAction SilentlyContinue
            Get-ChildItem -LiteralPath (Join-Path $skillPath "references") -File -ErrorAction SilentlyContinue
            Get-ChildItem -LiteralPath (Join-Path $skillPath "assets") -File -ErrorAction SilentlyContinue
            Get-ChildItem -LiteralPath (Join-Path $skillPath "scripts") -File -ErrorAction SilentlyContinue
        )
        Add-GroupFiles $items "ppt_master_skill" $skillFiles
    }
    return $items | Sort-Object group, absolute_path
}

$files = @(Collect-LockedFiles)
if ($Mode -eq "Write") {
    $manifest = [ordered]@{
        schema_version = 1
        status = "active"
        frozen_at = (Get-Date).ToString("o")
        rule = "Locked files may not be edited, renamed, moved, deleted, re-exported, or overwritten unless the user explicitly requests that change."
        exception_rule = "Create a backup and record the authorized scope before modifying a locked file."
        checksum_file = "TEMPLATE_LOCK.sha256"
        files = $files
    }
    $manifest | ConvertTo-Json -Depth 6 |
        Set-Content -LiteralPath $manifestPath -Encoding UTF8
    $lines = foreach ($item in $files) {
        $relative = if ($item.absolute_path.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
            [System.IO.Path]::GetRelativePath($root, $item.absolute_path).Replace("\", "/")
        }
        else {
            $item.absolute_path.Replace("\", "/")
        }
        "$($item.sha256.ToLowerInvariant())  $relative"
    }
    Set-Content -LiteralPath $checksumPath -Value $lines -Encoding UTF8
    [pscustomobject]@{
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
    if (-not (Test-Path -LiteralPath $item.absolute_path)) {
        $errors += [pscustomobject]@{ type = "missing"; path = $item.absolute_path }
        continue
    }
    $actual = (Get-FileHash -LiteralPath $item.absolute_path -Algorithm SHA256).Hash
    if ($actual -ne $item.sha256) {
        $errors += [pscustomobject]@{
            type = "changed"
            path = $item.absolute_path
            expected = $item.sha256
            actual = $actual
        }
    }
}
[pscustomobject]@{
    status = if ($errors.Count -eq 0) { "PASS" } else { "FAIL" }
    locked_files = @($stored.files).Count
    errors = $errors
} | ConvertTo-Json -Depth 5
if ($errors.Count -gt 0) {
    exit 1
}
