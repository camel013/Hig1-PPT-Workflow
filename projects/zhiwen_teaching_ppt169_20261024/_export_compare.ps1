$ErrorActionPreference = 'Stop'
$base = 'C:\Users\camel\Documents\ChatGPT\Hig1\projects\zhiwen_teaching_ppt169_20261024'
$out = Join-Path $base '_anim_compare'
New-Item -ItemType Directory -Force -Path (Join-Path $out 'before') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $out 'after') | Out-Null
$app = New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts = 1
try {
  foreach ($pair in @(
    @((Join-Path $base '蛭文_教学读本_95页_新讲解版.pptx'), (Join-Path $out 'before')),
    @((Join-Path $base '_grouped_anim_tmp.pptx'), (Join-Path $out 'after'))
  )) {
    $pres = $app.Presentations.Open($pair[0], $true, $false, $false)
    try {
      foreach ($slide in $pres.Slides) {
        $png = Join-Path $pair[1] ('slide-{0:D2}.png' -f $slide.SlideIndex)
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
