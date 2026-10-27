$ErrorActionPreference='Stop'
$path='C:\Users\camel\Documents\ChatGPT\Hig1\projects\zhiwen_teaching_ppt169_20261024\_anim_compare\after.pptx'
$app=New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts=1
$pres=$app.Presentations.Open($path,$true,$false,$false)
try {
  foreach($p in @(1,2,9,16,17,44,60,95)){
    $slide=$pres.Slides.Item($p); Write-Output ('--- PAGE '+$p+' ---')
    foreach($eff in $slide.TimeLine.MainSequence){
      Write-Output ('{0} | type={1} | dur={2:N2} | delay={3:N2}' -f $eff.Shape.Name,$eff.EffectType,$eff.Timing.Duration,$eff.Timing.TriggerDelayTime)
    }
  }
} finally {
  $pres.Close(); $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
