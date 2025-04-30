<#
.SYNOPSIS
    SharePoint Audit Log Collector with proper search creation throttling
.DESCRIPTION
    This script collects audit logs from SharePoint sites while managing Microsoft Graph API limits
    (10 concurrent searches maximum) with proper queuing and waiting logic.
.NOTES
    Version: 2.0
    Author: GeekByte
    Creation Date: 2024-04-02
#>

# Configuration Parameters
$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
#$ClientSecret = ""

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

# Constants
$MAX_CONCURRENT_SEARCHES = 10  # Microsoft Graph limit
$WAIT_TIME_SECONDS = 300       # 5 minutes wait time when limit reached
$RETRY_DELAY_SECONDS = 30      # Delay between status checks

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
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Failed to get access token: $_" -ForegroundColor Red
        exit
    }
}

# 2. Create Audit Search with throttling
function New-AuditSearch {
    param (
        [string]$AccessToken,
        [string]$Site,
        [string]$Operation,
        [ref]$ActiveSearchCount
    )

    $SearchParams = @{
        "displayName"           = "Audit_$($Site.Split('/')[-1])_$Operation_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
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
        # Check if we've reached the concurrent search limit
        if ($ActiveSearchCount.Value -ge $MAX_CONCURRENT_SEARCHES) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Waiting for available search slot (current: $($ActiveSearchCount.Value)/$MAX_CONCURRENT_SEARCHES)..." -ForegroundColor Yellow
            return $null
        }

        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Creating search for $Site - $Operation" -ForegroundColor Cyan
        $SearchQuery = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/security/auditLog/queries" `
                                        -Method Post `
                                        -Headers $Headers `
                                        -Body ($SearchParams | ConvertTo-Json -Depth 5)
        
        $ActiveSearchCount.Value++
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Search created (current active: $($ActiveSearchCount.Value))" -ForegroundColor Green
        return $SearchQuery.id
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Response.StatusCode -eq 503) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Rate limit hit when creating search. Waiting $WAIT_TIME_SECONDS seconds..." -ForegroundColor Red
            Start-Sleep -Seconds $WAIT_TIME_SECONDS
            return $null
        }
        else {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Failed to create audit search for $Site - $Operation : $_" -ForegroundColor Red
            return $null
        }
    }
}

# 3. Check Search Status with completion tracking
function Get-SearchStatus {
    param (
        [string]$AccessToken,
        [string]$SearchId,
        [ref]$ActiveSearchCount
    )

    $Headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }

    try {
        $SearchStatus = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/security/auditLog/queries/$SearchId" `
                                         -Method Get `
                                         -Headers $Headers
        
        if ($SearchStatus.status -eq "succeeded" -or $SearchStatus.status -eq "failed") {
            $ActiveSearchCount.Value--
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Search $SearchId completed with status: $($SearchStatus.status) (active searches now: $($ActiveSearchCount.Value))" -ForegroundColor Green
        }

        return $SearchStatus.status
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Response.StatusCode -eq 503) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Rate limit hit when checking status. Waiting $RETRY_DELAY_SECONDS seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds $RETRY_DELAY_SECONDS
            return "retry"
        }
        else {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Failed to get search status: $_" -ForegroundColor Red
            return "failed"
        }
    }
}

# 4. Retrieve Records with pagination
function Get-AuditRecords {
    param (
        [string]$AccessToken,
        [string]$SearchId
    )

    $AllRecords = @()
    $Uri = "https://graph.microsoft.com/beta/security/auditLog/queries/$SearchId/records?`$top=1000"
    $Headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }

    do {
        try {
            $Response = Invoke-RestMethod -Uri $Uri -Method Get -Headers $Headers
            $AllRecords += $Response.value
            
            if ($Response.'@odata.nextLink') {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Retrieved $($AllRecords.Count) records so far, more available..." -ForegroundColor Yellow
                $Uri = $Response.'@odata.nextLink'
                
                # Small delay between pages to avoid throttling
                Start-Sleep -Milliseconds 500
            }
            else {
                $Uri = $null
            }
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Response.StatusCode -eq 503) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Rate limit hit when retrieving records. Waiting $RETRY_DELAY_SECONDS seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $RETRY_DELAY_SECONDS
                continue
            }
            else {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Failed to retrieve records: $_" -ForegroundColor Red
                break
            }
        }
    } while ($Uri)

    return $AllRecords
}

# 5. Save to CSV with proper file naming
function Save-AuditToCsv {
    param (
        [array]$Records,
        [string]$Site,
        [string]$Operation
    )

    if ($Records.Count -eq 0) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] No records found for $Operation in $Site" -ForegroundColor Yellow
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
            auditData            = ($_ | Select-Object -ExpandProperty auditData | ConvertTo-Json -Depth 10 -Compress)
        }
    }

    try {
        if (Test-Path $FileName) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Appending to existing file: $FileName" -ForegroundColor Yellow
            $Report | Export-Csv $FileName -NoTypeInformation -Append -Force
        } 
        else {
            $Report | Export-Csv $FileName -NoTypeInformation -Force
        }
        
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Saved $($Records.Count) records to $FileName" -ForegroundColor Green
    }
    catch {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Failed to save CSV: $_" -ForegroundColor Red
    }
}

# Main Execution
function Main {
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting audit log collection process..." -ForegroundColor Cyan

    # Step 1: Authenticate
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Authenticating..." -ForegroundColor Yellow
    $AccessToken = Get-AccessToken

    # Step 2: Load existing searches from tracking files
    $ExistingSearches = @{}
    if (Test-Path "search_ids.json") {
        $ExistingSearches = Get-Content "search_ids.json" | ConvertFrom-Json -AsHashtable
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Loaded $($ExistingSearches.Count) existing searches from file." -ForegroundColor Yellow
    }

    $CompletedSearches = @{}
    if (Test-Path "completed_searches.json") {
        $CompletedSearches = Get-Content "completed_searches.json" | ConvertFrom-Json -AsHashtable
    }

    # Track active search count
    $ActiveSearchCount = 0

    # Step 3: Create searches with proper throttling
    $TotalSearches = $Sites.Count * $Operations.Count
    $CreatedSearches = 0

    foreach ($Site in $Sites) {
        foreach ($Operation in $Operations) {
            $Key = "${Site}_${Operation}"
            
            # Skip if already completed
            if ($CompletedSearches.ContainsKey($Key)) {
                $CreatedSearches++
                continue
            }

            # Create new search if needed
            if (-not $ExistingSearches.ContainsKey($Key)) {
                $SearchId = $null
                do {
                    $SearchId = New-AuditSearch -AccessToken $AccessToken -Site $Site -Operation $Operation -ActiveSearchCount ([ref]$ActiveSearchCount)
                    
                    if (-not $SearchId) {
                        # Wait if we hit the limit or got throttled
                        Start-Sleep -Seconds $RETRY_DELAY_SECONDS
                    }
                } while (-not $SearchId)

                if ($SearchId) {
                    $ExistingSearches[$Key] = @{
                        SearchId = $SearchId
                        CreatedTime = (Get-Date).ToUniversalTime().ToString("o")
                    }
                    $ExistingSearches | ConvertTo-Json | Out-File "search_ids.json"
                    $CreatedSearches++
                }
            }
            else {
                $CreatedSearches++
            }

            # Progress update
            Write-Progress -Activity "Creating searches" -Status "Progress: $CreatedSearches of $TotalSearches" -PercentComplete (($CreatedSearches / $TotalSearches) * 100)
        }
    }

    # Step 4: Process searches with proper concurrency management
    $ProcessedSearches = 0
    $SearchKeys = @($ExistingSearches.Keys | Where-Object { -not $CompletedSearches.ContainsKey($_) })

    foreach ($Key in $SearchKeys) {
        $SearchId = $ExistingSearches[$Key].SearchId
        $Status = $null

        do {
            $Status = Get-SearchStatus -AccessToken $AccessToken -SearchId $SearchId -ActiveSearchCount ([ref]$ActiveSearchCount)
            
            if ($Status -eq "retry") {
                Start-Sleep -Seconds $RETRY_DELAY_SECONDS
                continue
            }

            if ($Status -eq "succeeded") {
                # Extract site and operation from key
                $Parts = $Key -split "_", 2
                $Site = $Parts[0]
                $Operation = $Parts[1]

                # Retrieve records
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Retrieving records for $Site - $Operation" -ForegroundColor Cyan
                $Records = Get-AuditRecords -AccessToken $AccessToken -SearchId $SearchId

                # Save to CSV
                Save-AuditToCsv -Records $Records -Site $Site -Operation $Operation

                # Mark as completed
                $CompletedSearches[$Key] = @{
                    CompletedTime = (Get-Date).ToUniversalTime().ToString("o")
                    RecordCount = $Records.Count
                }
                $CompletedSearches | ConvertTo-Json | Out-File "completed_searches.json"
                
                $ProcessedSearches++
                Write-Progress -Activity "Processing searches" -Status "Progress: $ProcessedSearches of $($SearchKeys.Count)" -PercentComplete (($ProcessedSearches / $SearchKeys.Count) * 100)
            }
            elseif ($Status -eq "failed") {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Search $SearchId failed" -ForegroundColor Red
                $CompletedSearches[$Key] = @{
                    CompletedTime = (Get-Date).ToUniversalTime().ToString("o")
                    Status = "failed"
                }
                $CompletedSearches | ConvertTo-Json | Out-File "completed_searches.json"
                $ProcessedSearches++
            }
            else {
                # Still running, wait before checking again
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Search $SearchId status: $Status. Waiting $RETRY_DELAY_SECONDS seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $RETRY_DELAY_SECONDS
            }
        } while ($Status -ne "succeeded" -and $Status -ne "failed")
    }

    # Final cleanup
    if ($CompletedSearches.Count -eq $TotalSearches) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] All operations completed successfully!" -ForegroundColor Cyan
        if (Test-Path "search_ids.json") { Remove-Item "search_ids.json" }
        if (Test-Path "completed_searches.json") { Remove-Item "completed_searches.json" }
    } 
    else {
        $remaining = $TotalSearches - $CompletedSearches.Count
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $remaining searches remaining. Run the script again to continue." -ForegroundColor Yellow
    }
}

# Execute
Main
