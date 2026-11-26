# Recreates the interactive scheduled tasks eo-excel2pdf / eo-ppt2pdf by cloning eo-word2pdf (same principal, run in the interactive session).
schtasks /query /tn eo-word2pdf /xml | Out-File C:\eo-ref\t.xml -Encoding Unicode
$x = Get-Content C:\eo-ref\t.xml -Raw
$x.Replace('word2pdf.ps1','excel2pdf.ps1') | Set-Content C:\eo-ref\t1.xml -Encoding Unicode
$x.Replace('word2pdf.ps1','ppt2pdf.ps1') | Set-Content C:\eo-ref\t2.xml -Encoding Unicode
schtasks /create /tn eo-excel2pdf /xml C:\eo-ref\t1.xml /f
schtasks /create /tn eo-ppt2pdf /xml C:\eo-ref\t2.xml /f
Remove-Item C:\eo-ref\t*.xml
