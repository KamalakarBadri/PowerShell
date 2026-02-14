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




# SIMPLIFIED VERSION
$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ClientSecret = "nxH8Q~cnM2fkLf75~klgzsH9rzOC~~OgY.DEGaz"
$TenantName = "geekbyteonline.onmicrosoft.com"

# Get current date and time for filename
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$currentDateForReport = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Get Microsoft Graph token
Write-Host "Getting tokens..." -ForegroundColor Yellow


Write-Host "Graph token: OK" -ForegroundColor Green

# SharePoint token
$Certificate = Get-Item Cert:\CurrentUser\My\B799789F78628CAE56B4D0F380FD551EB754E0DB
$hash = [Convert]::ToBase64String($Certificate.GetCertHash())
$exp = [math]::Round((Get-Date).AddMinutes(10).ToFileTimeUtc()/10000000 - 11644473600)
$nbf = [math]::Round((Get-Date).ToFileTimeUtc()/10000000 - 11644473600)

$header = @{alg="RS256"; typ="JWT"; x5t=$hash -replace '\+','-' -replace '/','_' -replace '='}
$payload = @{aud="https://login.microsoftonline.com/$TenantName/oauth2/token"; exp=$exp; iss=$ClientId; jti=[guid]::NewGuid(); nbf=$nbf; sub=$ClientId}

$encodedHeader = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($header | ConvertTo-Json))) -replace '\+','-' -replace '/','_' -replace '='
$encodedPayload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json))) -replace '\+','-' -replace '/','_' -replace '='

$unsignedJWT = "$encodedHeader.$encodedPayload"
$privateKey = ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate))
$signature = [Convert]::ToBase64String($privateKey.SignData([Text.Encoding]::UTF8.GetBytes($unsignedJWT), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)) -replace '\+','-' -replace '/','_' -replace '='

$JWT = "$unsignedJWT.$signature"

$sharepointToken = (Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method Post -Body @{
    client_id=$ClientId; client_assertion=$JWT
    client_assertion_type="urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
    scope="https://geekbyteonline.sharepoint.com/.default"; grant_type="client_credentials"
} -ContentType "application/x-www-form-urlencoded").access_token
Write-Host "SharePoint token: OK" -ForegroundColor Green

$graphToken = (Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body @{
    client_id=$ClientId; client_assertion=$JWT
    client_assertion_type="urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
    scope="https://graph.microsoft.com/.default"; grant_type="client_credentials"
} -ContentType "application/x-www-form-urlencoded").access_token
