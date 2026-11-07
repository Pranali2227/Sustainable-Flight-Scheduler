$ErrorActionPreference = 'Stop'
$base = (Resolve-Path (Join-Path $PSScriptRoot 'Review_0_with_Review_1_Additions.docx')).Path
$pdf = Join-Path $PSScriptRoot 'Review_0_with_Review_1_Additions_QA.pdf'
$log = Join-Path $PSScriptRoot 'word_export.log'
Set-Content -Path $log -Value 'Starting Word'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
Add-Content -Path $log -Value 'Word created'
$doc = $word.Documents.Open($base, $false, $false)
Add-Content -Path $log -Value 'Document opened'
$doc.Fields.Update() | Out-Null
Add-Content -Path $log -Value 'Fields updated'
$doc.Save()
$doc.ExportAsFixedFormat($pdf, 17)
Add-Content -Path $log -Value 'PDF exported'
$doc.Close()
$word.Quit()
Add-Content -Path $log -Value 'Done'
