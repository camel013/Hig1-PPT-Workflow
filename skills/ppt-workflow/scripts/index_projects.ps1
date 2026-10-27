param(
    [string]$Project = "",
    [ValidateSet("Json", "Text")]
    [string]$Format = "Json"
)

$ErrorActionPreference = "Stop"
$WorkspaceRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..\..\..")).Path
$ProjectsRoot = Join-Path $WorkspaceRoot "projects"

function Get-RelativePath {
    param([string]$Path)
    [System.IO.Path]::GetRelativePath($WorkspaceRoot, $Path).Replace("\", "/")
}

function Get-ProjectSummary {
    param([System.IO.DirectoryInfo]$Directory)

    $entryExtensions = @(".py", ".mjs", ".js", ".ps1")
    $entrypoints = @(
        Get-ChildItem -LiteralPath $Directory.FullName -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in $entryExtensions } |
            Sort-Object Name |
            ForEach-Object { Get-RelativePath $_.FullName }
    )

    $sourceFiles = @()
    $sourceRoot = Join-Path $Directory.FullName "sources"
    if (Test-Path -LiteralPath $sourceRoot) {
        $sourceFiles = @(
            Get-ChildItem -LiteralPath $sourceRoot -Recurse -File -ErrorAction SilentlyContinue |
                Sort-Object FullName |
                ForEach-Object { Get-RelativePath $_.FullName }
        )
    }

    $assets = @()
    foreach ($assetName in @("images", "title_assets")) {
        $assetRoot = Join-Path $Directory.FullName $assetName
        if (Test-Path -LiteralPath $assetRoot) {
            $files = @(Get-ChildItem -LiteralPath $assetRoot -Recurse -File -ErrorAction SilentlyContinue)
            $measured = $files | Measure-Object Length -Sum
            $assets += [ordered]@{
                path = Get-RelativePath $assetRoot
                files = $files.Count
                bytes = [long]$measured.Sum
            }
        }
    }

    $deliverables = @()
    $deliveryRoot = Join-Path $Directory.FullName "deliverables"
    if (Test-Path -LiteralPath $deliveryRoot) {
        $deliverables = @(
            Get-ChildItem -LiteralPath $deliveryRoot -File -ErrorAction SilentlyContinue |
                Sort-Object Name |
                ForEach-Object {
                    [ordered]@{
                        path = Get-RelativePath $_.FullName
                        bytes = $_.Length
                    }
                }
        )
    }

    $readme = Join-Path $Directory.FullName "README.md"
    $spec = Join-Path $Directory.FullName "设计规范与内容大纲.md"
    return [ordered]@{
        project = $Directory.Name
        readme = if (Test-Path -LiteralPath $readme) { Get-RelativePath $readme } else { $null }
        design_spec = if (Test-Path -LiteralPath $spec) { Get-RelativePath $spec } else { $null }
        entrypoints = $entrypoints
        sources = $sourceFiles
        assets = $assets
        deliverables = $deliverables
    }
}

$projectDirectories = @(
    Get-ChildItem -LiteralPath $ProjectsRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { -not $Project -or $_.Name -eq $Project } |
        Sort-Object Name
)

if ($projectDirectories.Count -eq 0) {
    throw "No matching project found: $Project"
}

$index = @($projectDirectories | ForEach-Object { Get-ProjectSummary $_ })
if ($Format -eq "Text") {
    foreach ($item in $index) {
        $entryText = if ($item.entrypoints.Count) { $item.entrypoints -join ", " } else { "-" }
        $deliveryText = if ($item.deliverables.Count) { $item.deliverables.path -join ", " } else { "-" }
        Write-Output ("{0}`tspec={1}`tentry={2}`tdeliverables={3}" -f $item.project, $item.design_spec, $entryText, $deliveryText)
    }
}
else {
    $payload = if ($Project) { $index[0] } else { $index }
    $payload | ConvertTo-Json -Depth 6
}
