$ErrorActionPreference='Stop'
$path='C:\Users\camel\Documents\ChatGPT\Hig1\projects\zhiwen_teaching_ppt169_20261024\_anim_compare\after.pptx'
$app=New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts=1
$pres=$app.Presentations.Open($path,$true,$false,$false)
try {
  $rows=@()
  foreach($slide in $pres.Slides){
    $seq=$slide.TimeLine.MainSequence
    $names=@()
    foreach($eff in $seq){ $names += $eff.Shape.Name }
    $rows += [pscustomobject]@{Page=$slide.SlideIndex; Count=$seq.Count; Names=($names -join ' || ')}
  }
  $rows | Select-Object -First 20 | Format-Table -Wrap -AutoSize | Out-String -Width 240 | Write-Output
  Write-Output ('COUNTS:' + (($rows | ForEach-Object {$_.Count}) -join ','))
  foreach($p in @(1,2,9,16,17,44,60,95)){
    $r=$rows | Where-Object {$_.Page -eq $p}
    Write-Output ('PAGE '+$p+' '+$r.Count+' '+$r.Names)
  }
} finally {
  $pres.Close(); $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
