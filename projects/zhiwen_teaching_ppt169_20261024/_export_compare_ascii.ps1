$ErrorActionPreference = 'Stop'
$base = 'C:\Users\camel\Documents\ChatGPT\Hig1\projects\zhiwen_teaching_ppt169_20261024\_anim_compare'
foreach ($n in @('before','after')) { New-Item -ItemType Directory -Force -Path (Join-Path $base $n) | Out-Null }
$app = New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts = 1
try {
  foreach ($n in @('before','after')) {
    $pres = $app.Presentations.Open((Join-Path $base ($n+'.pptx')), $true, $false, $false)
    try {
      foreach ($slide in $pres.Slides) {
        $png = Join-Path (Join-Path $base $n) ('slide-{0:D2}.png' -f $slide.SlideIndex)
        $slide.Export($png, 'PNG', 1280, 720)
      }
    } finally { $pres.Close() }
  }
} finally {
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Output 'EXPORT_OK'
