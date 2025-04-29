$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"



# Date range (UTC)
$StartDate = "2025-01-10T00:00:00Z"
$EndDate = "2025-04-02T23:59:59Z"

# Sites and operations
$Sites = @(
       "https://geekbyteonline.sharepoint.com/sites/New365Site1",
    "https://geekbyteonline.sharepoint.com/sites/New365Site2",
    "https://geekbyteonline.sharepoint.com/sites/New365Site3",
    "https://geekbyteonline.sharepoint.com/sites/New365Site4",
    "https://geekbyteonline.sharepoint.com/sites/New365Site5"
)

$Operations = @("PageViewed", "FileAccessed", "FileDownloaded")
$ServiceFilter = "SharePoint"

# 1. Get Access Token
function Get-AccessToken {
    $AuthUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $Body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    try {
        $TokenResponse = Invoke-RestMethod -Uri $AuthUrl -Method Post -ContentType "application/x-www-form-urlencoded" -Body $Body
        return $TokenResponse.access_token
    }
    catch {
        Write-Host "Failed to get access token: $_" -ForegroundColor Red
        exit
    }
}

# 2. Create Audit Search
function New-AuditSearch {
    param (
        [string]$AccessToken,
        [string]$Site,
        [string]$Operation
    )

    $SearchParams = @{
        "displayName"           = "Audit_$($Site.Split('/')[-1])_$Operation_$(Get-Date -Format 'yyyyMMdd_HHmm')"
        "filterStartDateTime"   = $StartDate
        "filterEndDateTime"     = $EndDate
        "operationFilters"      = @($Operation)
        "serviceFilters"        = @($ServiceFilter)
        "objectIdFilters"       = @("$Site*")
    }

    $Headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }

    try {
        $SearchQuery = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/security/auditLog/queries" `
                                        -Method Post `
                                        -Headers $Headers `
                                        -Body ($SearchParams | ConvertTo-Json -Depth 5)
        return $SearchQuery.id
    }
    catch {
        Write-Host "Failed to create audit search for $Site - $Operation : $_" -ForegroundColor Red
        return $null
    }
}

# 3. Check Search Status
function Get-SearchStatus {
    param (
        [string]$AccessToken,
        [string]$SearchId
    )

    $Headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }

    try {
        $SearchStatus = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/security/auditLog/queries/$SearchId" `
                                         -Method Get `
                                         -Headers $Headers
        return $SearchStatus.status
    }
    catch {
        Write-Host "Failed to get search status: $_" -ForegroundColor Red
        return "failed"
    }
}

# 4. Retrieve Records
function Get-AuditRecords {
    param (
        [string]$AccessToken,
        [string]$SearchId
    )

    $AllRecords = @()
    $Uri = "https://graph.microsoft.com/beta/security/auditLog/queries/$SearchId/records?`$top=999"
    $Headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }

    do {
        try {
            $Response = Invoke-RestMethod -Uri $Uri -Method Get -Headers $Headers
            $AllRecords += $Response.value
            $Uri = $Response.'@odata.nextLink'
            
            if ($Uri) {
                Write-Host "Retrieved $($AllRecords.Count) records so far..." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "Failed to retrieve records: $_" -ForegroundColor Red
            break
        }
    } while ($Uri)

    return $AllRecords
}

# 5. Save to CSV
function Save-AuditToCsv {
    param (
        [array]$Records,
        [string]$Site,
        [string]$Operation
    )

    if ($Records.Count -eq 0) {
        Write-Host "No records found for $Operation in $Site" -ForegroundColor Yellow
        return
    }

    $SafeSite = ($Site -replace 'https?://|/', '_') -replace '[^a-zA-Z0-9_]', ''
    $FileName = "SharePointAudit_${SafeSite}_$Operation.csv"

    $Report = $Records | ForEach-Object {
        [PSCustomObject]@{
            id                   = $_.id
            createdDateTime      = $_.createdDateTime                    
            userPrincipalName    = $_.userPrincipalName                   
            operation            = $_.operation
            auditData            = ($_ | Select-Object -ExpandProperty auditData | ConvertTo-Json -Depth 10)
        }
    }

    # Check if file exists to append or create new
    if (Test-Path $FileName) {
        Write-Host "Appending to existing file: $FileName" -ForegroundColor Yellow
        $Report | Export-Csv $FileName -NoTypeInformation -Append
    } else {
        $Report | Export-Csv $FileName -NoTypeInformation -Force
    }
    
    Write-Host "Saved $($Records.Count) records to $FileName" -ForegroundColor Green
}

# 6. Get Existing Searches from Graph
function Get-ExistingGraphSearches {
    param (
        [string]$AccessToken
    )

    $Headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }

    try {
        $Response = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/security/auditLog/queries" `
                                     -Method Get `
                                     -Headers $Headers
        
        $Searches = @{}
        if ($Response.value) {
            foreach ($search in $Response.value) {
                $site = $search.objectIdFilters[0].TrimEnd('*')
                $operation = $search.operationFilters[0]
                $key = "${site}_${operation}"
                $Searches[$key] = $search.id
            }
        }
        return $Searches
    }
    catch {
        Write-Host "Failed to get existing searches from Graph: $_" -ForegroundColor Yellow
        return @{}
    }
}

# Main Execution
Write-Host "Starting audit log collection process..." -ForegroundColor Cyan

# Step 1: Authenticate
Write-Host "Authenticating..." -ForegroundColor Yellow
$AccessToken = Get-AccessToken

# Step 2: Load existing searches from file and Graph
$ExistingSearches = @{}
if (Test-Path "search_ids.json") {
    $ExistingSearches = Get-Content "search_ids.json" | ConvertFrom-Json -AsHashtable
    Write-Host "Loaded $($ExistingSearches.Count) existing searches from file." -ForegroundColor Yellow
}

# Get active searches from Graph
$ActiveSearches = Get-ExistingGraphSearches -AccessToken $AccessToken
Write-Host "Found $($ActiveSearches.Count) active searches in Microsoft Graph." -ForegroundColor Yellow

# Merge searches (prefer active ones)
$AllSearches = @{}
foreach ($key in $ExistingSearches.Keys) {
    $AllSearches[$key] = $ExistingSearches[$key]
}
foreach ($key in $ActiveSearches.Keys) {
    $AllSearches[$key] = $ActiveSearches[$key]
}

# Track completed searches separately
$CompletedSearches = @{}
if (Test-Path "completed_searches.json") {
    $CompletedSearches = Get-Content "completed_searches.json" | ConvertFrom-Json -AsHashtable
}

# Step 3: Create any missing searches with consistent 5-second delay
foreach ($Site in $Sites) {
    foreach ($Operation in $Operations) {
        $Key = "${Site}_${Operation}"
        if (-not $AllSearches.ContainsKey($Key) -and -not $CompletedSearches.ContainsKey($Key)) {
            Write-Host "Creating missing search for $Site - $Operation" -ForegroundColor Yellow
            $SearchId = New-AuditSearch -AccessToken $AccessToken -Site $Site -Operation $Operation
            if ($SearchId) {
                $AllSearches[$Key] = $SearchId
                
                # Save after each search creation
                $AllSearches | ConvertTo-Json | Out-File "search_ids.json"
                
                # Add 5-second delay after every search creation
                Write-Host "Waiting 5 seconds before next operation..." -ForegroundColor Yellow
                Start-Sleep -Seconds 5
            }
        }
    }
}

# Save current state of searches
$AllSearches | ConvertTo-Json | Out-File "search_ids.json"
Write-Host "Saved current search IDs to search_ids.json" -ForegroundColor Green

# Step 4: Process searches
# Clone the keys to avoid modification during enumeration
$SearchKeys = @($AllSearches.Keys)

foreach ($Key in $SearchKeys) {
    if ($CompletedSearches.ContainsKey($Key)) {
        continue
    }

    $SearchId = $AllSearches[$Key]
    $Status = Get-SearchStatus -AccessToken $AccessToken -SearchId $SearchId
    Write-Host "Search $SearchId ($Key) status: $Status" -ForegroundColor Yellow

    if ($Status -eq "succeeded") {
        # Extract site and operation from key
        $Parts = $Key -split "_", 2
        $Site = $Parts[0]
        $Operation = $Parts[1]

        # Retrieve records
        Write-Host "Retrieving records for $Site - $Operation" -ForegroundColor Yellow
        $Records = Get-AuditRecords -AccessToken $AccessToken -SearchId $SearchId

        # Save to CSV
        Save-AuditToCsv -Records $Records -Site $Site -Operation $Operation

        # Mark as completed (we won't delete since the API isn't working)
        $CompletedSearches[$Key] = $SearchId
        $CompletedSearches | ConvertTo-Json | Out-File "completed_searches.json"
        
        Write-Host "Completed processing for $Key" -ForegroundColor Green
    }
    elseif ($Status -eq "failed") {
        Write-Host "Search $SearchId failed" -ForegroundColor Red
        $CompletedSearches[$Key] = $SearchId
        $CompletedSearches | ConvertTo-Json | Out-File "completed_searches.json"
    }
}

# Check if all searches are completed
if ($CompletedSearches.Count -eq ($Sites.Count * $Operations.Count)) {
    Write-Host "All operations completed successfully!" -ForegroundColor Cyan
    if (Test-Path "search_ids.json") {
        Remove-Item "search_ids.json"
    }
    if (Test-Path "completed_searches.json") {
        Remove-Item "completed_searches.json"
    }
} else {
    $remaining = ($Sites.Count * $Operations.Count) - $CompletedSearches.Count
    Write-Host "$remaining searches remaining. You can run the script again later to continue." -ForegroundColor Yellow
    Write-Host "Not all searches completed yet. Waiting 5 minutes..." -ForegroundColor Yellow
    Start-Sleep -Seconds 300  # Wait 5 minutes before checking again
}
