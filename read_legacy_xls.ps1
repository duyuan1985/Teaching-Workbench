param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$connection = $null
try {
    $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$InputPath;Extended Properties='Excel 8.0;HDR=NO;IMEX=1'"
    $connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
    $connection.Open()
    $schema = $connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
    $sheets = @()
    foreach ($schemaRow in $schema.Rows) {
        $tableName = [string]$schemaRow.TABLE_NAME
        if (-not $tableName.EndsWith('$') -and -not $tableName.EndsWith("$'")) { continue }
        $queryName = $tableName.Trim("'").Replace("]", "]]" )
        $command = $connection.CreateCommand()
        $command.CommandText = "SELECT * FROM [$queryName]"
        $reader = $command.ExecuteReader()
        $rows = @()
        while ($reader.Read()) {
            $values = @()
            for ($index = 0; $index -lt $reader.FieldCount; $index++) {
                $value = $reader.GetValue($index)
                $values += $(if ($value -is [DBNull]) { "" } else { [string]$value })
            }
            $rows += ,$values
        }
        $reader.Close()
        $sheets += [ordered]@{ name = $tableName.TrimEnd("'", '$'); rows = $rows }
    }
    $json = ConvertTo-Json -InputObject $sheets -Depth 8 -Compress
    [System.IO.File]::WriteAllText($OutputPath, $json, (New-Object System.Text.UTF8Encoding($false)))
}
finally {
    if ($connection -ne $null) { $connection.Close(); $connection.Dispose() }
}
