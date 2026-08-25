param(
  [Parameter(Mandatory=$true)][string]$TemplatePath,
  [Parameter(Mandatory=$true)][string]$OutputPath,
  [Parameter(Mandatory=$true)][string]$DataPath
)
$ErrorActionPreference = 'Stop'
$data = Get-Content -LiteralPath $DataPath -Raw -Encoding UTF8 | ConvertFrom-Json
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath -Force
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
  $doc = $word.Documents.Open($OutputPath, $false, $false)
  function Set-Value($cell, $value, [double]$size=10.5) {
    $cell.Range.Text = [string]$value
    $cell.Range.Font.NameFarEast = '宋体'
    $cell.Range.Font.Name = '宋体'
    $cell.Range.Font.Size = $size
    $cell.VerticalAlignment = 1
  }
  $basic = $doc.Tables.Item(1)
  Set-Value $basic.Range.Cells.Item(2) $data.teacher
  Set-Value $basic.Range.Cells.Item(4) $data.department
  Set-Value $basic.Range.Cells.Item(6) $data.class_name
  Set-Value $basic.Range.Cells.Item(8) $data.course_name
  Set-Value $basic.Range.Cells.Item(10) $data.classroom
  Set-Value $basic.Range.Cells.Item(12) $data.word_time_label
  Set-Value $basic.Range.Cells.Item(16) $data.record 12
  $evaluation = $doc.Tables.Item(2)
  Set-Value $evaluation.Range.Cells.Item(5) $data.attendance
  Set-Value $evaluation.Range.Cells.Item(58) ("评语：" + $data.overall) 10
  Set-Value $evaluation.Range.Cells.Item(60) $data.suggestion 10
  Set-Value $evaluation.Range.Cells.Item(62) ("评价人：__________    日期：" + $data.date_label) 10
  $doc.Save()
  $doc.Close($false)
} finally {
  if($doc) { try { $doc.Close($false) } catch {} }
  $word.Quit()
}
