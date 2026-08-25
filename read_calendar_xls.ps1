$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $path = 'E:\工作\90-历史学期归档\2025-2026-2\课表、进程表、校历\2026年教学工作历(1).xls'
    $book = $excel.Workbooks.Open($path, $null, $true)
    foreach ($sheet in $book.Worksheets) {
        Write-Output ("SHEET {0} {1} {2}" -f $sheet.Name, $sheet.UsedRange.Rows.Count, $sheet.UsedRange.Columns.Count)
        for ($row = 1; $row -le $sheet.UsedRange.Rows.Count; $row++) {
            $values = @()
            for ($column = 1; $column -le $sheet.UsedRange.Columns.Count; $column++) {
                $value = $sheet.Cells.Item($row, $column).Text
                if ($value) { $values += $value }
            }
            if ($values.Count) { Write-Output ($values -join ' | ') }
        }
    }
    $book.Close($false)
}
finally {
    $excel.Quit()
}
