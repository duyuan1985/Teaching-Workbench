param(
  [Parameter(Mandatory=$true)][string]$InputPath,
  [Parameter(Mandatory=$true)][string]$OutputPath
)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
  $doc = $word.Documents.Open($InputPath, $false, $true)
  $parts = New-Object System.Collections.Generic.List[string]
  foreach($paragraph in $doc.Paragraphs) {
    $text = ($paragraph.Range.Text -replace '[\r\a]','').Trim()
    if($text) { $parts.Add($text) }
  }
  foreach($table in $doc.Tables) {
    foreach($cell in $table.Range.Cells) {
      $text = ($cell.Range.Text -replace '[\r\a]','').Trim()
      if($text) { $parts.Add($text) }
    }
  }
  [IO.File]::WriteAllText($OutputPath, ($parts -join "`n"), [Text.UTF8Encoding]::new($true))
  $doc.Close($false)
} finally {
  if($doc) { try { $doc.Close($false) } catch {} }
  $word.Quit()
}
