# export_data_to_excel.ps1

# Generates a report and exports it to XLSX format

# Database connection configuration (loaded from environment variables)
$server   = $env:DB_SERVER
$database = $env:DB_NAME
$user     = $env:DB_USER
$password = $env:DB_PASSWORD

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Output file names
$outputCSV  = "$timestamp`_data.csv"
$outputXLSX = "$timestamp`_data.xlsx"

# SQL query example (customize this according to your own database)
$query = @"
SELECT TOP 10
    id AS ID,
    name AS Name,
    email AS Email,
    created_at AS CreatedDate
FROM users
"@

if ($password -eq "") {
    $connectionString = "Data Source=$server;Initial Catalog=$database;User ID=$user;Password=;TrustServerCertificate=True;"
} else {
    $connectionString = "Data Source=$server;Initial Catalog=$database;User ID=$user;Password=$password;TrustServerCertificate=True;"
}

Write-Host "Connecting to database '$database'..."

try {
    $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $conn.Open()

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $query
    $cmd.CommandTimeout = 300

    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    $dataTable = New-Object System.Data.DataTable
    $adapter.Fill($dataTable) | Out-Null

    $conn.Close()
    Write-Host "Query executed successfully. Records retrieved: $($dataTable.Rows.Count)"
}
catch {
    Write-Error "Failed to execute query: $_"
    exit 1
}

try {
    if (Test-Path $outputCSV) { Remove-Item $outputCSV -Force }
    $dataTable | Export-Csv -Path $outputCSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    (Get-Content $outputCSV) -replace '"', '' | Set-Content $outputCSV -Encoding UTF8
    Write-Host "CSV file generated: $(Resolve-Path $outputCSV)"
}
catch {
    Write-Error "Failed to generate CSV file: $_"
    exit 1
}

if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..."
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

Write-Host "Converting CSV to XLSX..."

try {
    # Sheet 1: Data
    Import-Csv -Path $outputCSV -Delimiter ";" | Export-Excel -Path $outputXLSX -WorksheetName "Data"

    # Sheet 2: Example fixed data
    $sheet2Data = @(
        [PSCustomObject]@{
            "Product GUID"   = "00000000-0000-0000-0000-000000000001"
            "Product Name"   = "Basic Plan"
            "Product Code"   = "PROD001"
        },
        [PSCustomObject]@{
            "Product GUID"   = "00000000-0000-0000-0000-000000000002"
            "Product Name"   = "Advanced Plan"
            "Product Code"   = "PROD002"
        }
    )

    $sheet2Data | Export-Excel -Path $outputXLSX -WorksheetName "Reference" -Append

    Write-Host "XLSX file generated: $(Resolve-Path $outputXLSX)"
}
catch {
    Write-Warning "Failed to convert CSV to XLSX. Details: $_"
    Write-Warning "CSV is available at: $(Resolve-Path $outputCSV)"
    exit 1
}

try {
    Remove-Item $outputCSV -Force
    Write-Host "Temporary CSV file deleted."
}
catch {
    Write-Warning "Failed to delete CSV file. $_"
}
