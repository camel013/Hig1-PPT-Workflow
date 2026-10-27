param(
    [string]$DeckName = "meeting_4p",
    [int]$Pages = 4,
    [string]$RunRoot = "",
    [string]$QualityReference = "C:\Users\camel\Documents\ChatGPT\Hig1\projects\duoyuhua_ppt169_20261021\多余的话_瑞士国际主义版式.pptx",
    [string]$DesignSpec = "C:\Users\camel\Documents\ChatGPT\Hig1\projects\duoyuhua_ppt169_20261021\设计规范与内容大纲.md"
)

$ErrorActionPreference = "Stop"

$WorkspaceRoot = "C:\Users\camel\Documents\ChatGPT\Hig1"
$Skill = "C:\Users\camel\.codex\skills\ppt-master-skill"
$SkillMd = Join-Path $Skill "SKILL.md"
$VenvPython = Join-Path $Skill ".venv\Scripts\python.exe"
$Vendor = Join-Path $Skill "vendor\ppt-master"
$Edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

if (-not $RunRoot) {
    $RunRoot = Join-Path $WorkspaceRoot ("_ppt_runs\" + $DeckName + "_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
}

foreach ($dir in @($RunRoot, (Join-Path $RunRoot "svg_output"), (Join-Path $RunRoot "svg_final"), (Join-Path $RunRoot "preview"), (Join-Path $RunRoot "notes"), (Join-Path $RunRoot "build"))) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$checks = [ordered]@{
    workspace = Test-Path -LiteralPath $WorkspaceRoot
    skill_md = Test-Path -LiteralPath $SkillMd
    venv_python = Test-Path -LiteralPath $VenvPython
    vendor = Test-Path -LiteralPath $Vendor
    edge = Test-Path -LiteralPath $Edge
    quality_reference = Test-Path -LiteralPath $QualityReference
    design_spec = Test-Path -LiteralPath $DesignSpec
}

$fontDirs = @(
    "C:\Windows\Fonts",
    (Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Fonts"),
    "C:\Program Files\Microsoft Office\root\vfs\Fonts\private",
    "C:\Program Files (x86)\Microsoft Office\root\vfs\Fonts\private"
)

$fontFiles = @()
foreach ($dir in $fontDirs) {
    if (Test-Path -LiteralPath $dir) {
        $fontFiles += Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match 'simkai|STZHONG|STZhong|中宋|kai' } |
            Select-Object -ExpandProperty FullName
    }
}

$registryFonts = @()
foreach ($key in @("HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", "HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")) {
    try {
        $props = (Get-ItemProperty -LiteralPath $key -ErrorAction Stop).PSObject.Properties |
            Where-Object { $_.Name -match '中宋|Zhongsong|楷体|KaiTi' } |
            ForEach-Object { "$($key)|$($_.Name)|$($_.Value)" }
        $registryFonts += $props
    }
    catch {
    }
}

$pythonInfo = $null
if ($checks.venv_python) {
    try {
        $pythonInfo = & $VenvPython -c "import sys, json; import pptx; import PIL; print(json.dumps({'python': sys.version.split()[0], 'pptx': pptx.__version__, 'pillow': PIL.__version__}, ensure_ascii=False))"
    }
    catch {
        $pythonInfo = $null
    }
}

$manifest = [ordered]@{
    schema_version = 1
    generated_at = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss zzz")
    deck = [ordered]@{
        name = $DeckName
        pages = $Pages
        canvas = "1280x720"
        run_root = $RunRoot
    }
    paths = [ordered]@{
        workspace = $WorkspaceRoot
        skill = $Skill
        skill_md = $SkillMd
        venv_python = $VenvPython
        vendor = $Vendor
        edge = $Edge
        quality_reference = $QualityReference
        design_spec = $DesignSpec
        pptx = (Join-Path $RunRoot "deck.pptx")
        contact_sheet = (Join-Path $RunRoot "contact_sheet.png")
    }
    checks = $checks
    python = if ($pythonInfo) { $pythonInfo | ConvertFrom-Json } else { $null }
    fonts = [ordered]@{
        files = $fontFiles
        registry = $registryFonts
        kai_present = [bool]($fontFiles | Where-Object { $_ -match 'simkai' } | Select-Object -First 1)
        song_present = [bool]($fontFiles | Where-Object { $_ -match 'STZHONG|STZhong|中宋' } | Select-Object -First 1)
    }
    design_tokens = [ordered]@{
        paper = "#F4F1E9"
        ink = "#232323"
        muted = "#6E6A64"
        rule = "#D3CEC3"
        terracotta = "#8E5750"
        slate = "#5A707E"
        ochre = "#B49360"
        sage = "#7D8D7B"
        pale = "#F7F3EA"
        title_font = "楷体"
        body_font = "华文中宋"
        margins = "80/1200"
        header_line = 96
        footer_line = 664
        page_number = 692
    }
    review_policy = [ordered]@{
        contact_sheet_quick_screen = $true
        per_page_full_size = $true
        max_full_size_views_per_page = 1
        extra_view_only_if_issue = $true
    }
    budget = [ordered]@{
        hard_model_calls = 25
        a_preflight_max = 2
        b_content_and_design_min = 6
        c_build_max = 2
        d_review_and_fix_max = 7
        design_iteration_reserve = 5
    }
}

$manifestPath = Join-Path $RunRoot "manifest.json"
$manifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

[pscustomobject]@{
    manifest = $manifestPath
    run_root = $RunRoot
    checks = $checks
    kai_present = $manifest.fonts.kai_present
    song_present = $manifest.fonts.song_present
    python = $manifest.python
} | ConvertTo-Json -Depth 6 -Compress
