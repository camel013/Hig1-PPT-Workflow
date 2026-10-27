$ErrorActionPreference='Stop'
$base='C:\Users\camel\Documents\ChatGPT\Hig1\projects\zhiwen_teaching_ppt169_20261024\_anim_compare'
$out=Join-Path $base 'final'; New-Item -ItemType Directory -Force -Path $out | Out-Null
$app=New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts=1
$pres=$app.Presentations.Open((Join-Path $base 'final.pptx'),$true,$false,$false)
try {
  $counts=@()
  foreach($slide in $pres.Slides){
    $counts += $slide.TimeLine.MainSequence.Count
    $png=Join-Path $out ('slide-{0:D2}.png' -f $slide.SlideIndex)
    $slide.Export($png,'PNG',1280,720)
  }
  Write-Output ('FINAL_COUNTS:' + ($counts -join ','))
} finally {
  $pres.Close();$app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres)|Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
