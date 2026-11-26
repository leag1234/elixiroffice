# excel2pdf.ps1 - export every C:\eo-ref\inx\*.xlsx (whole workbook, file's own page setup/print areas) to C:\eo-ref\outx\<name>.pdf + .json via desktop Excel (COM).
# Run inside the interactive session (scheduled task eo-excel2pdf). Writes outx\run.log and outx\done.marker.
$ErrorActionPreference = 'Continue'
$in = 'C:\eo-ref\inx'; $out = 'C:\eo-ref\outx'
New-Item -ItemType Directory -Force $out | Out-Null
$log = Join-Path $out 'run.log'; $marker = Join-Path $out 'done.marker'
Remove-Item $marker -ErrorAction SilentlyContinue
function L($m) { ("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $m) | Add-Content -Encoding UTF8 $log }
Remove-Item $log -ErrorAction SilentlyContinue
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
$files = Get-ChildItem $in -Filter *.xlsx | Sort-Object Name
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
    $x = $null
    try {
      $x = New-Object -ComObject Excel.Application
      $x.Visible = $false; $x.DisplayAlerts = $false; $x.EnableEvents = $false
      $x.AutomationSecurity = 3          # msoAutomationSecurityForceDisable: never run macros
      $x.AskToUpdateLinks = $false
      $r.excel_version = $x.Version; $r.excel_build = $x.Build
      # Open(Filename, UpdateLinks=0 never, ReadOnly=true, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended=true, ..., AddToMru=false)
      $m = [Type]::Missing
      $wb = $x.Workbooks.Open($src, 0, $true, $m, $m, $m, $true, $m, $m, $m, $false, $m, $false)
      $r.sheets = $wb.Sheets.Count
      $vis = 0; foreach ($s in $wb.Sheets) { if ($s.Visible -eq -1) { $vis++ } }
      $r.visible_sheets = $vis
      # ExportAsFixedFormat(Type=0 PDF, Filename, Quality=0 standard, IncludeDocProperties, IgnorePrintAreas=false, From, To, OpenAfterPublish=false) on the Workbook = all sheets
      $wb.ExportAsFixedFormat(0, $pdf, 0, $true, $false, $m, $m, $false)
      $wb.Close($false)
      $r.ok = (Test-Path $pdf)
    } catch { $r.error = $_.Exception.Message }
    finally { try { if ($x) { $x.Quit() } } catch {} }
    $r | ConvertTo-Json | Set-Content -Encoding UTF8 $json
  }
  if (-not (Wait-Job $job -Timeout 120)) {
    L "[$name] TIMEOUT 120s, killing EXCEL"
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
    Stop-Job $job; $res = @{ file = $f.Name; ok = $false; error = 'timeout 120s' }
    $res | ConvertTo-Json | Set-Content -Encoding UTF8 $json
  }
  Remove-Job $job -Force
  try {
    $r = Get-Content $json -Raw -Encoding UTF8 | ConvertFrom-Json
    $r | Add-Member -NotePropertyName seconds -NotePropertyValue ([math]::Round($sw.Elapsed.TotalSeconds,1)) -Force
    $r | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 $json
    L "[$name] done ok=$($r.ok) sheets=$($r.sheets) $($r.error)"
  } catch { L "[$name] json merge failed: $($_.Exception.Message)" }
  Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
Start-Sleep 2
L "finished; EXCEL left: $((Get-Process EXCEL -ErrorAction SilentlyContinue | measure).Count)"
'done' | Set-Content $marker
