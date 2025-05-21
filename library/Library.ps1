# Load the SharePoint access token from tokens.json
$tokens = Get-Content -Path "tokens.json" | ConvertFrom-Json
$sharepointToken = $tokens.sharepoint_token
$siteUrl = "https://geekbyteonline.sharepoint.com/sites/Newwww2"

# Fixed SharePoint groups (modify these as needed)
$sharepointGroups = @(
    "Newwww Owners"
    
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

function Add-UserToLibrary {
    param (
        [string]$token,
        [string]$siteUrl,
        [string]$listId,
        [string]$userEmail,
        [int]$roleDefId
    )
    
    try {
        $userPrincipal = Ensure-Principal -token $token -siteUrl $siteUrl -principalIdentifier $userEmail
        Add-RoleAssignment -token $token -siteUrl $siteUrl -listId $listId -principalId $userPrincipal.Id -roleDefId $roleDefId
        Write-Host "Added user $userEmail with access" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Could not add user $userEmail : $_" -ForegroundColor Yellow
        return $false
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
        $breakResponse = Invoke-RestMethod -Uri "$siteUrl/_api/web/lists('$listId')/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)" -Method Post -Headers $headers
        Write-Host "Broke role inheritance successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Error breaking role inheritance: $_" -ForegroundColor Red
        exit
    }
    
    # 3. Get role definitions
    try {
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
		"DHCP Users"
    )
    
    foreach ($groupName in $domainGroups) {
        try {
            $groupPrincipal = Ensure-Principal -token $sharepointToken -siteUrl $siteUrl -principalIdentifier $groupName -isGroup $true
            Add-RoleAssignment -token $sharepointToken -siteUrl $siteUrl -listId $listId -principalId $groupPrincipal.Id -roleDefId $roles['Contribute']
            Write-Host "Added domain group '$($groupPrincipal.Title)' with Contribute access" -ForegroundColor Green
        }
        catch {
            Write-Host "Could not add domain group $groupName : $_" -ForegroundColor Yellow
        }
    }
    
    # 5. Add SharePoint groups with Full Control
    foreach ($groupName in $sharepointGroups) {
        try {
            $spGroupId = Get-GroupId -token $sharepointToken -siteUrl $siteUrl -groupName $groupName
            Add-RoleAssignment -token $sharepointToken -siteUrl $siteUrl -listId $listId -principalId $spGroupId -roleDefId $roles['Full Control']
            Write-Host "Added SharePoint group '$groupName' with Full Control access" -ForegroundColor Green
        }
        catch {
            Write-Host "Could not add SharePoint group $groupName : $_" -ForegroundColor Yellow
        }
    }
	
    # 6. Add SharePoint groups with Contribute
		$groupName = "Newwww Members"
        try {
            $spGroupId = Get-GroupId -token $sharepointToken -siteUrl $siteUrl -groupName $groupName
            Add-RoleAssignment -token $sharepointToken -siteUrl $siteUrl -listId $listId -principalId $spGroupId -roleDefId $roles['Contribute']
            Write-Host "Added SharePoint group '$groupName' with Full Control access" -ForegroundColor Green
        }
        catch {
            Write-Host "Could not add SharePoint group $groupName : $_" -ForegroundColor Yellow
        }
   	
    
    # # 6. Optional user provisioning
    # $usersInput = Read-Host "Enter user emails to add (comma separated, or leave blank)"
    # if ($usersInput) {
        # $userEmails = $usersInput -split "," | ForEach-Object { $_.Trim() }
        # foreach ($email in $userEmails) {
            # Add-UserToLibrary -token $sharepointToken -siteUrl $siteUrl -listId $listId -userEmail $email -roleDefId $roles['Contribute']
        # }
    # }
    
    Write-Host "`nDocument library setup complete!" -ForegroundColor Green
    Write-Host "Library Name: $libraryName"
    Write-Host "Library ID: $listId"
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
