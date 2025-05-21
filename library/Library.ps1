<#
.SYNOPSIS
    Creates a SharePoint document library with custom permissions and generates required tokens automatically
.DESCRIPTION
    This script:
    1. Generates Microsoft Graph and SharePoint access tokens using certificate authentication
    2. Creates a new document library with the name FAX_(ProjectName)_Site
    3. Breaks permission inheritance
    4. Adds domain groups with Contribute access
    5. Adds SharePoint groups with specific permissions (Owners: Full Control, Admins: Contribute)
    6. Optionally adds individual users
#>

#region Token Generation
$TenantName = "sampleten.onmicrosoft.com"  
$AppId = "" 
$Certificate = Get-Item Cert:\CurrentUser\My\ 
$ScopeGraph = "https://graph.microsoft.com/.default"
$ScopeSharePoint = "https://sampleten.sharepoint.com/.default"

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
    $sharepointToken = Get-AccessToken -Resource "https://sampleten.sharepoint.com" -Scope $ScopeSharePoint
    Write-Host "Successfully generated access tokens" -ForegroundColor Green
}
catch {
    Write-Host "Error generating access tokens: $_" -ForegroundColor Red
    exit
}
#endregion

#region Document Library Creation
$siteUrl = "https://sampleten.sharepoint.com/sites/Newwww2"

# SharePoint groups with their required access levels
$sharepointGroups = @(
    @{ Name = "FAX Site Owners"; Role = "Full Control" },
    @{ Name = "FAX Site Admins"; Role = "Contribute" }
)

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

function Get-GroupId {
    param (
        [string]$token,
        [string]$siteUrl,
        [string]$groupName
    )
    
    $endpoint = "$siteUrl/_api/web/sitegroups/getbyname('$groupName')"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Accept" = "application/json;odata=verbose"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $endpoint -Method Get -Headers $headers
        return $response.d.Id
    }
    catch {
        Write-Host "Error getting group ID for $groupName : $_" -ForegroundColor Red
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

# Main script execution
try {
    # Get project name from user
    $projectName = Read-Host "Enter the project name"
    $libraryName = "FAX_${projectName}_Site"
    
    # Get request digest
    $requestDigest = Get-RequestDigest -token $sharepointToken -siteUrl $siteUrl
    
    # Create headers
    $headers = @{
        "Authorization" = "Bearer $sharepointToken"
        "Accept" = "application/json;odata=verbose"
        "Content-Type" = "application/json;odata=verbose"
        "X-RequestDigest" = $requestDigest
    }
    
    # 1. Create document library
    $libraryData = @{
        "__metadata" = @{"type" = "SP.List"}
        "AllowContentTypes" = $true
        "BaseTemplate" = 101
        "ContentTypesEnabled" = $true
        "Description" = "Document library for $projectName project"
        "Title" = $libraryName
    } | ConvertTo-Json
    
    try {
        Write-Host "Creating document library '$libraryName'..." -ForegroundColor Cyan
        $createResponse = Invoke-RestMethod -Uri "$siteUrl/_api/web/lists" -Method Post -Headers $headers -Body $libraryData
        $listId = $createResponse.d.Id
        Write-Host "Created library '$libraryName' (ID: $listId)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error creating library: $_" -ForegroundColor Red
        exit
    }
    
    # 2. Break role inheritance
    try {
        Write-Host "Breaking permission inheritance..." -ForegroundColor Cyan
        $breakResponse = Invoke-RestMethod -Uri "$siteUrl/_api/web/lists('$listId')/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)" -Method Post -Headers $headers
        Write-Host "Broke role inheritance successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Error breaking role inheritance: $_" -ForegroundColor Red
        exit
    }
    
    # 3. Get role definitions
    try {
        Write-Host "Getting permission levels..." -ForegroundColor Cyan
        $rolesResponse = Invoke-RestMethod -Uri "$siteUrl/_api/web/roledefinitions" -Method Get -Headers @{
            "Authorization" = "Bearer $sharepointToken"
            "Accept" = "application/json;odata=verbose"
        }
        $roles = @{}
        $rolesResponse.d.results | ForEach-Object {
            $roles[$_.Name] = $_.Id
        }
    }
    catch {
        Write-Host "Error getting role definitions: $_" -ForegroundColor Red
        exit
    }
    
    # 4. Add domain groups with Contribute access
    $domainGroups = @(
        "FAX_${projectName}_AZG",
        "FAX_${projectName}_ZZZ"
    )
    
    foreach ($groupName in $domainGroups) {
        try {
            Write-Host "Adding domain group '$groupName' with Contribute access..." -ForegroundColor Cyan
            $groupPrincipal = Ensure-Principal -token $sharepointToken -siteUrl $siteUrl -principalIdentifier $groupName -isGroup $true
            Add-RoleAssignment -token $sharepointToken -siteUrl $siteUrl -listId $listId -principalId $groupPrincipal.Id -roleDefId $roles['Contribute']
            Write-Host "Added domain group '$($groupPrincipal.Title)' with Contribute access" -ForegroundColor Green
        }
        catch {
            Write-Host "Could not add domain group $groupName : $_" -ForegroundColor Yellow
        }
    }
    
    # 5. Add SharePoint groups with specific access levels
    foreach ($group in $sharepointGroups) {
        try {
            Write-Host "Adding SharePoint group '$($group.Name)' with $($group.Role) access..." -ForegroundColor Cyan
            $spGroupId = Get-GroupId -token $sharepointToken -siteUrl $siteUrl -groupName $group.Name
            Add-RoleAssignment -token $sharepointToken -siteUrl $siteUrl -listId $listId -principalId $spGroupId -roleDefId $roles[$group.Role]
            Write-Host "Added SharePoint group '$($group.Name)' with $($group.Role) access" -ForegroundColor Green
        }
        catch {
            Write-Host "Could not add SharePoint group $($group.Name) : $_" -ForegroundColor Yellow
        }
    }
    
    # 6. Optional user provisioning
    $usersInput = Read-Host "Enter user emails to add (comma separated, or leave blank)"
    if ($usersInput) {
        $userEmails = $usersInput -split "," | ForEach-Object { $_.Trim() }
        foreach ($email in $userEmails) {
            try {
                Write-Host "Adding user '$email' with Contribute access..." -ForegroundColor Cyan
                $userPrincipal = Ensure-Principal -token $sharepointToken -siteUrl $siteUrl -principalIdentifier $email
                Add-RoleAssignment -token $sharepointToken -siteUrl $siteUrl -listId $listId -principalId $userPrincipal.Id -roleDefId $roles['Contribute']
                Write-Host "Added user $email with Contribute access" -ForegroundColor Green
            }
            catch {
                Write-Host "Could not add user $email : $_" -ForegroundColor Yellow
            }
        }
    }
    
    # Final output
    Write-Host "`nDocument library setup complete!" -ForegroundColor Green
    Write-Host "Library Name: $libraryName"
    Write-Host "Library ID: $listId"
    Write-Host "`nPermission Summary:"
    Write-Host "- Domain Groups (Contribute): $($domainGroups -join ', ')"
    foreach ($group in $sharepointGroups) {
        Write-Host "- $($group.Name) ($($group.Role))"
    }
    if ($usersInput) {
        Write-Host "- Additional Users (Contribute): $usersInput"
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