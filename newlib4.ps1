<#
.SYNOPSIS
    Fetches SharePoint document libraries starting with "SPO" and adds "newGroup" with Contribute permission
.DESCRIPTION
    This script:
    1. Generates Microsoft Graph and SharePoint access tokens using certificate authentication
    2. Fetches all document libraries with names starting with "SPO"
    3. Adds "newGroup" with Contribute permission to each library
#>

#region Token Generation
$TenantName = "geekbyteonline.onmicrosoft.com"  
$AppId = "73efa35d-6188-42d4-b258-838a977eb149" 
$Certificate = Get-Item Cert:\CurrentUser\My\B799789F78628CAE56B4D0F380FD551EB754E0DB  
$ScopeGraph = "https://graph.microsoft.com/.default"
$ScopeSharePoint = "https://geekbyteonline.sharepoint.com/.default"

function Get-AccessToken {
    param (
        [string]$Resource,
        [string]$Scope
    )
    
    # Create base64 hash of certificate  
    $CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())  

    # Create JWT timestamps
    $StartDate = (Get-Date "1970-01-01T00:00:00Z").ToUniversalTime()  
    $JWTExpiration = [math]::Round((New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds, 0)
    $NotBefore = [math]::Round((New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds, 0)

    # Create JWT header  
    $JWTHeader = @{  
        alg = "RS256"  
        typ = "JWT"  
        x5t = $CertificateBase64Hash -replace '\+','-' -replace '/','_' -replace '='  
    }  

    # Create JWT payload
    $JWTPayload = @{  
        aud = "https://login.microsoftonline.com/$TenantName/oauth2/token"  
        exp = $JWTExpiration  
        iss = $AppId  
        jti = [guid]::NewGuid()  
        nbf = $NotBefore  
        sub = $AppId  
    }  

    # Convert to JWT
    $JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))  
    $EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)  
    $JWTPayloadToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))  
    $EncodedPayload = [System.Convert]::ToBase64String($JWTPayloadToByte)  
    $JWT = $EncodedHeader + "." + $EncodedPayload  

    # Sign the JWT
    $PrivateKey = ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate))  
    $Signature = [Convert]::ToBase64String(  
        $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)  
    ) -replace '\+','-' -replace '/','_' -replace '='  
    $JWT = $JWT + "." + $Signature  

    # Request token
    $Body = @{  
        client_id = $AppId  
        client_assertion = $JWT  
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"  
        scope = $Scope  
        grant_type = "client_credentials"  
    }  

    $Url = "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token"  
    $Response = Invoke-RestMethod -Uri $Url -Method Post -Body $Body -ContentType 'application/x-www-form-urlencoded'
    return $Response.access_token
}

# Get both tokens
try {
    Write-Host "Generating access tokens..." -ForegroundColor Cyan
    $graphToken = Get-AccessToken -Resource "https://graph.microsoft.com" -Scope $ScopeGraph
    $sharepointToken = Get-AccessToken -Resource "https://geekbyteonline.sharepoint.com" -Scope $ScopeSharePoint
    Write-Host "Successfully generated access tokens" -ForegroundColor Green
}
catch {
    Write-Host "Error generating access tokens: $_" -ForegroundColor Red
    exit
}
#endregion

#region Document Library Processing
$siteUrl = "https://geekbyteonline.sharepoint.com/sites/Newwww2"
$groupName = "New SC"

function Get-RequestDigest {
    param (
        [string]$token,
        [string]$siteUrl
    )
    
    $contextInfoUrl = "$siteUrl/_api/contextinfo"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Accept" = "application/json;odata=verbose"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $contextInfoUrl -Method Post -Headers $headers
        return $response.d.GetContextWebInformation.FormDigestValue
    }
    catch {
        Write-Host "Error getting request digest: $_" -ForegroundColor Red
        exit
    }
}

function Ensure-Principal {
    param (
        [string]$token,
        [string]$siteUrl,
        [string]$principalIdentifier,
        [bool]$isGroup = $false
    )
    
    $endpoint = "$siteUrl/_api/web/ensureuser"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Accept" = "application/json;odata=verbose"
        "Content-Type" = "application/json;odata=verbose"
        "X-RequestDigest" = (Get-RequestDigest -token $token -siteUrl $siteUrl)
    }
    
    if ($isGroup -and $principalIdentifier -notmatch "@" -and $principalIdentifier -notmatch "\|") {
        # Try different group name formats
        $formats = @(
            $principalIdentifier,
            "c:0t.c|tenant|$principalIdentifier"
        )
        
        foreach ($format in $formats) {
            try {
                $body = @{
                    'logonName' = $format
                } | ConvertTo-Json
                
                $response = Invoke-RestMethod -Uri $endpoint -Method Post -Headers $headers -Body $body
                return $response.d
            }
            catch {
                continue
            }
        }
        throw "Could not resolve group: $principalIdentifier"
    }
    else {
        $loginName = if ($principalIdentifier -match "@" -or $principalIdentifier -match "\|") {
            $principalIdentifier
        } else {
            "i:0#.f|membership|$principalIdentifier"
        }
        
        $body = @{
            'logonName' = $loginName
        } | ConvertTo-Json
        
        try {
            $response = Invoke-RestMethod -Uri $endpoint -Method Post -Headers $headers -Body $body
            return $response.d
        }
        catch {
            Write-Host "Error ensuring principal $principalIdentifier : $_" -ForegroundColor Red
            throw
        }
    }
}

function Get-RoleDefinitions {
    param (
        [string]$token,
        [string]$siteUrl
    )
    
    $endpoint = "$siteUrl/_api/web/roledefinitions"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Accept" = "application/json;odata=verbose"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $endpoint -Method Get -Headers $headers
        $roles = @{}
        $response.d.results | ForEach-Object {
            $roles[$_.Name] = $_.Id
        }
        return $roles
    }
    catch {
        Write-Host "Error getting role definitions: $_" -ForegroundColor Red
        throw
    }
}

function Add-RoleAssignment {
    param (
        [string]$token,
        [string]$siteUrl,
        [string]$listId,
        [string]$principalId,
        [int]$roleDefId
    )
    
    $endpoint = "$siteUrl/_api/web/lists('$listId')/roleassignments/addroleassignment(principalid=$principalId,roledefid=$roleDefId)"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Accept" = "application/json;odata=verbose"
        "X-RequestDigest" = (Get-RequestDigest -token $token -siteUrl $siteUrl)
    }
    
    try {
        $response = Invoke-RestMethod -Uri $endpoint -Method Post -Headers $headers
        return $response
    }
    catch {
        Write-Host "Error adding role assignment: $_" -ForegroundColor Red
        throw
    }
}

function Get-SPOLibraries {
    param (
        [string]$token,
        [string]$siteUrl
    )
    
    $endpoint = "$siteUrl/_api/web/lists?`$filter=BaseTemplate eq 101 and startswith(Title,'FAX')"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Accept" = "application/json;odata=verbose"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $endpoint -Method Get -Headers $headers
        return $response.d.results
    }
    catch {
        Write-Host "Error fetching SPO libraries: $_" -ForegroundColor Red
        throw
    }
}

# Main script execution
try {
    Write-Host "Starting SPO library processing..." -ForegroundColor Cyan
    
    # Get request digest
    $requestDigest = Get-RequestDigest -token $sharepointToken -siteUrl $siteUrl
    
    # 1. Get all document libraries starting with "SPO"
    Write-Host "Fetching document libraries starting with 'SPO'..." -ForegroundColor Cyan
    $spoLibraries = Get-SPOLibraries -token $sharepointToken -siteUrl $siteUrl
    
    if ($spoLibraries.Count -eq 0) {
        Write-Host "No document libraries found starting with 'SPO'" -ForegroundColor Yellow
        exit
    }
    
    Write-Host "Found $($spoLibraries.Count) SPO libraries:" -ForegroundColor Green
    $spoLibraries | ForEach-Object { Write-Host "  - $($_.Title) (ID: $($_.Id))" -ForegroundColor White }
    
    # 2. Get role definitions
    Write-Host "`nGetting permission levels..." -ForegroundColor Cyan
    $roles = Get-RoleDefinitions -token $sharepointToken -siteUrl $siteUrl
    
    if (-not $roles.ContainsKey('Contribute')) {
        Write-Host "Error: 'Contribute' role definition not found" -ForegroundColor Red
        exit
    }
    
    $contributeRoleId = $roles['Contribute']
    
    # 3. Ensure the group exists in SharePoint
    Write-Host "Ensuring group '$groupName' exists in SharePoint..." -ForegroundColor Cyan
    try {
        $groupPrincipal = Ensure-Principal -token $sharepointToken -siteUrl $siteUrl -principalIdentifier $groupName -isGroup $true
        Write-Host "Group '$groupName' resolved successfully (ID: $($groupPrincipal.Id))" -ForegroundColor Green
    }
    catch {
        Write-Host "Error: Could not resolve group '$groupName'. Please ensure the group exists. Error: $_" -ForegroundColor Red
        exit
    }
    
    # 4. Process each SPO library
    $successCount = 0
    $errorCount = 0
    
    foreach ($library in $spoLibraries) {
        Write-Host "`nProcessing library: $($library.Title)..." -ForegroundColor Cyan
        
        try {
            # Check if library already has unique permissions (optional step)
            # If you want to ensure inheritance is broken first, uncomment the following section:
            <#
            Write-Host "  Breaking permission inheritance..." -ForegroundColor Yellow
            $breakHeaders = @{
                "Authorization" = "Bearer $sharepointToken"
                "Accept" = "application/json;odata=verbose"
                "X-RequestDigest" = $requestDigest
            }
            Invoke-RestMethod -Uri "$siteUrl/_api/web/lists('$($library.Id)')/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)" -Method Post -Headers $breakHeaders
            Write-Host "  Permission inheritance broken" -ForegroundColor Green
            #>
            
            # Add group with Contribute permission
            Write-Host "  Adding '$groupName' with Contribute access..." -ForegroundColor Yellow
            Add-RoleAssignment -token $sharepointToken -siteUrl $siteUrl -listId $library.Id -principalId $groupPrincipal.Id -roleDefId $contributeRoleId
            Write-Host "  Successfully added '$groupName' to $($library.Title)" -ForegroundColor Green
            $successCount++
        }
        catch {
            Write-Host "  Error processing library $($library.Title): $_" -ForegroundColor Red
            $errorCount++
        }
    }
    
    # Final summary
    Write-Host "`n" + "="*50 -ForegroundColor Cyan
    Write-Host "PROCESSING SUMMARY" -ForegroundColor Cyan
    Write-Host "="*50 -ForegroundColor Cyan
    Write-Host "Total SPO libraries found: $($spoLibraries.Count)" -ForegroundColor White
    Write-Host "Successfully processed: $successCount" -ForegroundColor Green
    Write-Host "Errors: $errorCount" -ForegroundColor Red
    Write-Host "Group added: $groupName" -ForegroundColor White
    Write-Host "Permission level: Contribute" -ForegroundColor White
    
    if ($errorCount -eq 0) {
        Write-Host "`nAll SPO libraries have been successfully updated!" -ForegroundColor Green
    } else {
        Write-Host "`nCompleted with errors. Please check the log above." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Error in main script execution: $_" -ForegroundColor Red
    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseText = $reader.ReadToEnd()
        Write-Host "Response: $responseText" -ForegroundColor Yellow
    }
}
#endregion
