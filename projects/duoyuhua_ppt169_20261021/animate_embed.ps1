# 使用前：先把 build_pptx.py 生成的成品复制为 _deck_build.pptx
$ErrorActionPreference = 'Stop'
$dir = 'C:\Users\camel\Documents\ChatGPT\Hig1\projects\duoyuhua_ppt169_20261021'
$src = Join-Path $dir '_deck_build.pptx'
$subset = Join-Path $dir '_deck_subset.pptx'
$tmp = Join-Path $dir '_deck_final.pptx'
$outDir = Join-Path $dir 'pp_check'
$app = New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts = 1
$pres = $app.Presentations.Open($src, $false, $false, $false)
$counts = @()
foreach ($slide in $pres.Slides) {
    $shapes = @($slide.Shapes)
    $idx = 0
    foreach ($shape in $shapes) {
        if ($shape.Name -like 'bg*') { continue }
        $eff = $slide.TimeLine.MainSequence.AddEffect($shape, 10, 0, 2, -1)
        $eff.Timing.Duration = 0.4
        $eff.Timing.TriggerDelayTime = 0.06 * $idx
        $idx++
    }
    $counts += $idx
    $png = Join-Path $outDir ('s{0:D2}.png' -f $slide.SlideIndex)
    $slide.Export($png, 'PNG', 1280, 720)
}
try { $pres.EmbedTrueTypeFonts = -1 } catch { }
$pres.SaveAs($subset, 24, -1)
$pres.Close()
$app.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Start-Sleep -Seconds 2
$py = 'C:\Users\camel\.codex\skills\ppt-master-skill\.venv\Scripts\python.exe'
& $py -X utf8 (Join-Path $dir 'embed_full_fonts.py') $subset $tmp
if ($LASTEXITCODE -ne 0) { throw 'embed_full_fonts.py failed' }
Remove-Item -LiteralPath $subset -Force
Write-Output ('ANIM:' + ($counts -join ','))
Write-Output ('SIZE:' + (Get-Item -LiteralPath $tmp).Length)



