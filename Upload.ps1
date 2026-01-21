# ==============================================================================
# SIMPLE VERSION: Upload Single CSV to SharePoint List
# ==============================================================================

# SET YOUR VALUES HERE
$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ClientSecret = "t"
$SiteId = "geekbyteonline.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$ListId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# CSV File - extract tenant name from filename
$CsvFile = "C:\Data\Tenant1.csv"
$TenantName = [System.IO.Path]::GetFileNameWithoutExtension($CsvFile)

Clear-Host
Write-Host "Uploading $TenantName to SharePoint..." -ForegroundColor Cyan

# 1. Get Token
Write-Host "Getting access token..." -ForegroundColor Yellow

try {
    $tokenResponse = Invoke-RestMethod `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Method Post `
        -Body @{
            client_id = $ClientId
            client_secret = $ClientSecret
            scope = "https://graph.microsoft.com/.default"
            grant_type = "client_credentials"
        } `
        -ContentType "application/x-www-form-urlencoded"

    $token = $tokenResponse.access_token
    Write-Host "Token obtained" -ForegroundColor Green
} catch {
    Write-Host "Failed to get token: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# 2. Read CSV
Write-Host "Reading CSV file..." -ForegroundColor Yellow

if (-not (Test-Path $CsvFile)) {
    Write-Host "CSV file not found: $CsvFile" -ForegroundColor Red
    exit
}

try {
    $csv = Import-Csv $CsvFile
    Write-Host "CSV Data:" -ForegroundColor Gray
    $csv | Format-Table
} catch {
    Write-Host "Failed to read CSV: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# 3. Prepare data
$rowData = @{}
$rowData["Title"] = $TenantName  # Tenant name as Title

foreach ($row in $csv) {
    $columnName = $row.ColumnType
    $columnValue = $row.Count
    $rowData[$columnName] = $columnValue
}

# 4. Upload to SharePoint
Write-Host "Uploading to SharePoint..." -ForegroundColor Yellow

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type" = "application/json"
}

$body = @{ fields = $rowData } | ConvertTo-Json
$url = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$ListId/items"

try {
    $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
    Write-Host "Success! Item created with ID: $($response.id)" -ForegroundColor Green
    
    # Show what was uploaded
    Write-Host "`nUploaded data for $TenantName:" -ForegroundColor Cyan
    foreach ($key in $rowData.Keys) {
        Write-Host "  $key : $($rowData[$key])" -ForegroundColor Gray
    }
} catch {
    Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nDone!" -ForegroundColor Cyan
