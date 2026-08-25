$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open('D:\Downloads\教学安排表20260816115742.xls', $null, $true)
    foreach ($sheet in $book.Worksheets) {
        Write-Output ("SHEET {0} {1} {2}" -f $sheet.Name, $sheet.UsedRange.Rows.Count, $sheet.UsedRange.Columns.Count)
        $maxRow = [Math]::Min($sheet.UsedRange.Rows.Count, 100)
        for ($row = 1; $row -le $maxRow; $row++) {
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
