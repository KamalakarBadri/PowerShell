# SIMPLIFIED VERSION
$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ClientSecret = "nxHEGaz7"

# Get token
$tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
}

$response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
Write-host " $response "
$accessToken = $response.access_token

# Get all groups with pagination
$headers = @{ "Authorization" = "Bearer $accessToken" }
$allGroups = @()
$url = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,mail,visibility,createdDateTime,expirationDateTime&`$top=999"

while ($url) {
    $result = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
    $allGroups += $result.value
    $url = $result.'@odata.nextLink'
}

# Export to CSV
$csvData = foreach ($group in $allGroups) {
    [PSCustomObject]@{
        Name = $group.displayName
        Email = $group.mail
        Visibility = $group.visibility
        Created = if ($group.createdDateTime) { 
            [DateTime]::Parse($group.createdDateTime).ToString("yyyy-MM-dd") 
        } else { "" }
        Expires = if ($group.expirationDateTime) { 
            [DateTime]::Parse($group.expirationDateTime).ToString("yyyy-MM-dd") 
        } else { "" }
    }
}

$csvData | Export-Csv -Path "Groups.csv" -NoTypeInformation
Write-Host "Exported $($allGroups.Count) groups to Groups.csv" -ForegroundColor Green
