# ppt2pdf.ps1 - export every C:\eo-ref\inp\*.pptx to C:\eo-ref\outp\<name>.pdf + .json via desktop PowerPoint (COM): one slide per page, no notes, hidden slides included.
# Run inside the interactive session (scheduled task eo-ppt2pdf). Writes outp\run.log and outp\done.marker.
$ErrorActionPreference = 'Continue'
$in = 'C:\eo-ref\inp'; $out = 'C:\eo-ref\outp'
New-Item -ItemType Directory -Force $out | Out-Null
$log = Join-Path $out 'run.log'; $marker = Join-Path $out 'done.marker'
Remove-Item $marker -ErrorAction SilentlyContinue
function L($m) { ("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $m) | Add-Content -Encoding UTF8 $log }
Remove-Item $log -ErrorAction SilentlyContinue
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
$files = Get-ChildItem $in -Filter *.pptx | Sort-Object Name
L "start: $($files.Count) files"
foreach ($f in $files) {
  $name = [IO.Path]::GetFileNameWithoutExtension($f.Name)
  $pdf = Join-Path $out "$name.pdf"; $json = Join-Path $out "$name.json"
  Remove-Item $pdf, $json -ErrorAction SilentlyContinue
  L "[$name] begin"
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $job = Start-Job -ArgumentList $f.FullName, $pdf, $json -ScriptBlock {
    param($src, $pdf, $json)
    $r = @{ file = [IO.Path]::GetFileName($src); ok = $false }
    $p = $null
    try {
      $p = New-Object -ComObject PowerPoint.Application
      $p.DisplayAlerts = 1                # ppAlertsNone
      $p.AutomationSecurity = 3           # msoAutomationSecurityForceDisable
      $r.ppt_version = $p.Version; $r.ppt_build = $p.Build
      # Open(FileName, ReadOnly=true, Untitled=false, WithWindow=false)
      $d = $p.Presentations.Open($src, -1, 0, 0)
      $r.slides = $d.Slides.Count
      $h = 0; foreach ($s in $d.Slides) { if ($s.SlideShowTransition.Hidden -eq -1) { $h++ } }
      $r.hidden_slides = $h
      $m = [Type]::Missing
      $d.SaveAs($pdf, 32, 0)   # ppSaveAsPDF; ExportAsFixedFormat fails through the PowerShell COM adapter (enum conversion)
      $d.Close()
      $r.ok = (Test-Path $pdf)
    } catch { $r.error = $_.Exception.Message }
    finally { try { if ($p) { $p.Quit() } } catch {} }
    $r | ConvertTo-Json | Set-Content -Encoding UTF8 $json
  }
  if (-not (Wait-Job $job -Timeout 120)) {
    L "[$name] TIMEOUT 120s, killing POWERPNT"
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
    Stop-Job $job; $res = @{ file = $f.Name; ok = $false; error = 'timeout 120s' }
    $res | ConvertTo-Json | Set-Content -Encoding UTF8 $json
  }
  Remove-Job $job -Force
  try {
    $r = Get-Content $json -Raw -Encoding UTF8 | ConvertFrom-Json
    $r | Add-Member -NotePropertyName seconds -NotePropertyValue ([math]::Round($sw.Elapsed.TotalSeconds,1)) -Force
    $r | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 $json
    L "[$name] done ok=$($r.ok) slides=$($r.slides) hidden=$($r.hidden_slides) $($r.error)"
  } catch { L "[$name] json merge failed: $($_.Exception.Message)" }
  Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
}
Start-Sleep 2
L "finished; POWERPNT left: $((Get-Process POWERPNT -ErrorAction SilentlyContinue | measure).Count)"
'done' | Set-Content $marker
