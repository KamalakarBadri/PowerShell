<#
.SYNOPSIS
    SharePoint Audit Log Collector with Advanced Throttling and Error Handling
.DESCRIPTION
    Collects SharePoint audit logs with:
    - Automatic 5-minute pause after every 10 search creations
    - Special handling for token expiration and server errors
    - Comprehensive summary reporting
.NOTES
    Version: 3.2
    Author: GeekByte
    Creation Date: 2024-04-05
#>

# Configuration Parameters
$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ClientSecret = "CyG8

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

# Throttling Constants
$MAX_CONCURRENT_SEARCHES = 10
$SEARCH_BATCH_WAIT_MINUTES = 5
$RETRY_DELAY_SECONDS = 5
$TOKEN_EXPIRY_WAIT = 30
$PAGE_RETRIEVAL_DELAY = 500 # milliseconds

# Output Configuration
$REPORT_FOLDER = "AuditReports_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $REPORT_FOLDER -Force | Out-Null

# Global Variables
$script:AccessToken = $null
$script:LastTokenRefresh = [datetime]::MinValue

#region Helper Functions
function Log-Info($message) {
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] INFO: $message" -ForegroundColor Cyan
}

function Log-Success($message) {
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] SUCCESS: $message" -ForegroundColor Green
}

function Log-Warning($message) {
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARNING: $message" -ForegroundColor Yellow
}

function Log-Error($message) {
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $message" -ForegroundColor Red
}

function Test-TokenExpiry {
    # Refresh token if it's about to expire (or if we don't have one)
    if (-not $script:AccessToken -or $script:LastTokenRefresh -lt (Get-Date).AddMinutes(-55)) {
        Log-Info "Refreshing access token..."
        $script:AccessToken = Get-AccessToken
        $script:LastTokenRefresh = Get-Date
    }
}
#endregion

#region Core Functions
function Get-AccessToken {
    param([bool]$Retry = $false)
    
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
        if ($Retry) {
            Log-Error "Failed to get access token after retry: $_"
            exit
        }
        
        Log-Warning "Token request failed. Waiting $TOKEN_EXPIRY_WAIT seconds before retry..."
        Start-Sleep -Seconds $TOKEN_EXPIRY_WAIT
        return Get-AccessToken -Retry $true
    }
}

function New-AuditSearch {
    param (
        [string]$Site,
        [string]$Operation,
        [ref]$ActiveSearchCount,
        [ref]$SearchCreationCounter
    )

    Test-TokenExpiry

    $SearchName = "Audit_$($Site.Split('/')[-1])_$Operation_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $SearchParams = @{
        "displayName"           = $SearchName
        "filterStartDateTime"   = $StartDate
        "filterEndDateTime"     = $EndDate
        "operationFilters"      = @($Operation)
        "serviceFilters"        = @($ServiceFilter)
        "objectIdFilters"       = @("$Site*")
    }

    $Headers = @{
        "Authorization" = "Bearer $script:AccessToken"
        "Content-Type" = "application/json"
    }

    try {
        # Check concurrent search limit
        if ($ActiveSearchCount.Value -ge $MAX_CONCURRENT_SEARCHES) {
            Log-Warning "Concurrent search limit reached ($($ActiveSearchCount.Value)/$MAX_CONCURRENT_SEARCHES). Waiting..."
            return $null
        }

        # Check if we've created a batch of 10 searches
        if ($SearchCreationCounter.Value -gt 0 -and $SearchCreationCounter.Value % 10 -eq 0) {
            $waitTime = $SEARCH_BATCH_WAIT_MINUTES * 60
            Log-Info "Created 10 searches. Waiting $SEARCH_BATCH_WAIT_MINUTES minutes before creating more..."
            Start-Sleep -Seconds $waitTime
        }

        Log-Info "Creating search for $($Site.Split('/')[-1]) - $Operation"
        $SearchQuery = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/security/auditLog/queries" `
                                        -Method Post `
                                        -Headers $Headers `
                                        -Body ($SearchParams | ConvertTo-Json -Depth 5)
        
        $ActiveSearchCount.Value++
        $SearchCreationCounter.Value++
        Log-Success "Search created (ID: $($SearchQuery.id), active: $($ActiveSearchCount.Value))"
        
        return @{
            SearchId = $SearchQuery.id
            SearchName = $SearchName
            Site = $Site
            Operation = $Operation
            CreatedTime = (Get-date).ToUniversalTime().ToString("o")
            Status = "created"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        
        # Handle specific error cases
        if ($errorMsg -like "*internal server error*" -or $errorMsg -like "*tenant*") {
            Log-Warning "Server error encountered. Waiting $RETRY_DELAY_SECONDS seconds before retry..."
            Start-Sleep -Seconds $RETRY_DELAY_SECONDS
            return $null
        }
        elseif ($errorMsg -like "*token is expired*" -or $errorMsg -like "*authentication failed*") {
            Log-Warning "Token expired. Refreshing..."
            $script:AccessToken = Get-AccessToken
            return $null
        }
        elseif ($_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Response.StatusCode -eq 503) {
            $waitTime = $SEARCH_BATCH_WAIT_MINUTES * 60
            Log-Warning "Rate limit hit. Waiting $SEARCH_BATCH_WAIT_MINUTES minutes..."
            Start-Sleep -Seconds $waitTime
            return $null
        }
        else {
            Log-Error "Failed to create search: $_"
            return $null
        }
    }
}

function Get-SearchStatus {
    param (
        [hashtable]$SearchInfo,
        [ref]$ActiveSearchCount
    )

    Test-TokenExpiry

    $Headers = @{
        "Authorization" = "Bearer $script:AccessToken"
        "Content-Type" = "application/json"
    }

    try {
        $SearchStatus = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/security/auditLog/queries/$($SearchInfo.SearchId)" `
                                         -Method Get `
                                         -Headers $Headers
        
        if ($SearchStatus.status -in @("succeeded", "failed")) {
            $ActiveSearchCount.Value--
            $SearchInfo.Status = $SearchStatus.status
            $SearchInfo.CompletedTime = (Get-Date).ToUniversalTime().ToString("o")
            
            if ($SearchStatus.status -eq "succeeded") {
                Log-Success "Search completed: $($SearchInfo.SearchName)"
            } else {
                Log-Warning "Search failed: $($SearchInfo.SearchName)"
            }
        }

        return $SearchStatus.status
    }
    catch {
        $errorMsg = $_.Exception.Message
        
        if ($errorMsg -like "*token is expired*" -or $errorMsg -like "*authentication failed*") {
            Log-Warning "Token expired during status check. Refreshing..."
            $script:AccessToken = Get-AccessToken
            return "token_expired"
        }
        elseif ($errorMsg -like "*internal server error*" -or $errorMsg -like "*tenant*") {
            Log-Warning "Server error during status check. Waiting $RETRY_DELAY_SECONDS seconds..."
            Start-Sleep -Seconds $RETRY_DELAY_SECONDS
            return "retry"
        }
        else {
            Log-Error "Failed to get search status: $_"
            return "failed"
        }
    }
}

function Get-AuditRecords {
    param (
        [string]$SearchId
    )

    Test-TokenExpiry

    $AllRecords = @()
    $Uri = "https://graph.microsoft.com/beta/security/auditLog/queries/$SearchId/records?`$top=1000"
    $Headers = @{
        "Authorization" = "Bearer $script:AccessToken"
        "Content-Type" = "application/json"
    }

    do {
        try {
            $Response = Invoke-RestMethod -Uri $Uri -Method Get -Headers $Headers
            $AllRecords += $Response.value
            
            if ($Response.'@odata.nextLink') {
                Log-Info "Retrieved $($AllRecords.Count) records, more available..."
                $Uri = $Response.'@odata.nextLink'
                Start-Sleep -Milliseconds $PAGE_RETRIEVAL_DELAY
            } else {
                $Uri = $null
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            
            if ($errorMsg -like "*token is expired*" -or $errorMsg -like "*authentication failed*") {
                Log-Warning "Token expired during record retrieval. Refreshing..."
                $script:AccessToken = Get-AccessToken
                continue
            }
            elseif ($errorMsg -like "*internal server error*" -or $errorMsg -like "*tenant*") {
                Log-Warning "Server error during record retrieval. Waiting $RETRY_DELAY_SECONDS seconds..."
                Start-Sleep -Seconds $RETRY_DELAY_SECONDS
                continue
            }
            else {
                Log-Error "Failed to retrieve records: $_"
                break
            }
        }
    } while ($Uri)

    return $AllRecords
}

function Export-AuditData {
    param (
        [array]$Records,
        [hashtable]$SearchInfo
    )

    $SiteName = $SearchInfo.Site.Split('/')[-1]
    $FileName = "$REPORT_FOLDER\$($SiteName)_$($SearchInfo.Operation).csv"
    $SummaryInfo = @{
        Site = $SearchInfo.Site
        SiteName = $SiteName
        Operation = $SearchInfo.Operation
        RecordCount = $Records.Count
        FileName = $FileName
        SearchId = $SearchInfo.SearchId
        CreatedTime = $SearchInfo.CreatedTime
        CompletedTime = $SearchInfo.CompletedTime
    }

    if ($Records.Count -eq 0) {
        Log-Warning "No records found for $($SearchInfo.Operation) in $SiteName"
        return $SummaryInfo
    }

    $Report = $Records | ForEach-Object {
        [PSCustomObject]@{
            Timestamp          = $_.createdDateTime                    
            User               = $_.userPrincipalName                   
            Operation          = $_.operation
            Site               = $SearchInfo.Site
            Details            = ($_ | Select-Object -ExpandProperty auditData | ConvertTo-Json -Depth 5 -Compress)
        }
    }

    try {
        $Report | Export-Csv $FileName -NoTypeInformation -Force
        Log-Success "Saved $($Records.Count) records to $FileName"
        return $SummaryInfo
    }
    catch {
        Log-Error "Failed to save CSV: $_"
        return $SummaryInfo
    }
}

function Generate-SummaryReport {
    param (
        [array]$AllSummaries
    )

    $SummaryFile = "$REPORT_FOLDER\Audit_Summary_Report.csv"
    $DetailedSummaryFile = "$REPORT_FOLDER\Audit_Detailed_Summary.csv"

    # Generate summary by site and operation
    $SummaryBySiteOperation = $AllSummaries | Group-Object SiteName, Operation | ForEach-Object {
        [PSCustomObject]@{
            Site = $_.Group[0].Site
            SiteName = $_.Group[0].SiteName
            Operation = $_.Group[0].Operation
            TotalRecords = ($_.Group | Measure-Object -Property RecordCount -Sum).Sum
            Files = $_.Group.FileName -join "`n"
            SearchIDs = $_.Group.SearchId -join "`n"
        }
    }

    $SummaryBySiteOperation | Export-Csv $SummaryFile -NoTypeInformation -Force
    Log-Success "Generated summary report: $SummaryFile"

    # Generate detailed summary
    $AllSummaries | Select-Object Site, SiteName, Operation, RecordCount, FileName, SearchId, CreatedTime, CompletedTime |
        Export-Csv $DetailedSummaryFile -NoTypeInformation -Force
    Log-Success "Generated detailed summary: $DetailedSummaryFile"
}
#endregion

# Main Execution
function Main {
    Log-Info "Starting audit log collection process"
    $script:AccessToken = Get-AccessToken
    $script:LastTokenRefresh = Get-Date

    # Initialize counters
    $ActiveSearchCount = 0
    $SearchCreationCounter = 0
    $TotalSearches = $Sites.Count * $Operations.Count

    # Load or initialize tracking
    $SearchTrackingFile = "$REPORT_FOLDER\search_tracking.json"
    $CompletedTrackingFile = "$REPORT_FOLDER\completed_tracking.json"
    
    $AllSearches = @{}
    if (Test-Path $SearchTrackingFile) {
        $AllSearches = Get-Content $SearchTrackingFile | ConvertFrom-Json -AsHashtable
        Log-Info "Loaded $($AllSearches.Count) existing searches from tracking file"
        $ActiveSearchCount = ($AllSearches.Values | Where-Object { $_.Status -eq "created" }).Count
    }

    $CompletedSummaries = @()
    if (Test-Path $CompletedTrackingFile) {
        $CompletedSummaries = Get-Content $CompletedTrackingFile | ConvertFrom-Json
        Log-Info "Loaded $($CompletedSummaries.Count) completed searches"
    }

    # Phase 1: Create needed searches with batch control
    $CreatedSearches = 0
    foreach ($Site in $Sites) {
        foreach ($Operation in $Operations) {
            $Key = "$($Site.Split('/')[-1])_$Operation"
            
            # Skip if already completed
            if ($CompletedSummaries | Where-Object { "$($_.SiteName)_$($_.Operation)" -eq $Key }) {
                $CreatedSearches++
                continue
            }

            # Create new search if needed
            if (-not $AllSearches.ContainsKey($Key)) {
                $SearchInfo = $null
                do {
                    $SearchInfo = New-AuditSearch -Site $Site -Operation $Operation `
                                                -ActiveSearchCount ([ref]$ActiveSearchCount) `
                                                -SearchCreationCounter ([ref]$SearchCreationCounter)
                    
                    if (-not $SearchInfo) {
                        Start-Sleep -Seconds $RETRY_DELAY_SECONDS
                    }
                } while (-not $SearchInfo)

                if ($SearchInfo) {
                    $AllSearches[$Key] = $SearchInfo
                    $AllSearches | ConvertTo-Json -Depth 5 | Out-File $SearchTrackingFile -Force
                    $CreatedSearches++
                }
            } else {
                $CreatedSearches++
            }

            Write-Progress -Activity "Creating searches" -Status "$CreatedSearches of $TotalSearches" -PercentComplete ($CreatedSearches/$TotalSearches*100)
        }
    }

    # Phase 2: Process searches
    $ProcessedSearches = 0
    $SearchKeys = @($AllSearches.Keys | Where-Object {
        -not ($CompletedSummaries | Where-Object { "$($_.SiteName)_$($_.Operation)" -eq $_ })
    })

    foreach ($Key in $SearchKeys) {
        $SearchInfo = $AllSearches[$Key]
        $Status = $null

        do {
            $Status = Get-SearchStatus -SearchInfo $SearchInfo -ActiveSearchCount ([ref]$ActiveSearchCount)
            
            if ($Status -eq "retry" -or $Status -eq "token_expired") {
                Start-Sleep -Seconds $RETRY_DELAY_SECONDS
                continue
            }

            if ($Status -eq "succeeded") {
                $Records = Get-AuditRecords -SearchId $SearchInfo.SearchId
                $Summary = Export-AuditData -Records $Records -SearchInfo $SearchInfo
                $CompletedSummaries += $Summary
                $CompletedSummaries | ConvertTo-Json -Depth 5 | Out-File $CompletedTrackingFile -Force
                
                $ProcessedSearches++
                Write-Progress -Activity "Processing searches" -Status "$ProcessedSearches of $($SearchKeys.Count)" -PercentComplete ($ProcessedSearches/$SearchKeys.Count*100)
            }
            elseif ($Status -eq "failed") {
                $CompletedSummaries += @{
                    Site = $SearchInfo.Site
                    SiteName = $SearchInfo.Site.Split('/')[-1]
                    Operation = $SearchInfo.Operation
                    RecordCount = 0
                    FileName = ""
                    SearchId = $SearchInfo.SearchId
                    CreatedTime = $SearchInfo.CreatedTime
                    CompletedTime = $SearchInfo.CompletedTime
                    Status = "failed"
                }
                $CompletedSummaries | ConvertTo-Json -Depth 5 | Out-File $CompletedTrackingFile -Force
                $ProcessedSearches++
            }
            else {
                Log-Info "Search $($SearchInfo.SearchName) status: $Status. Waiting $RETRY_DELAY_SECONDS seconds..."
                Start-Sleep -Seconds $RETRY_DELAY_SECONDS
            }
        } while ($Status -notin @("succeeded", "failed"))
    }

    # Generate final reports
    if ($CompletedSummaries) {
        Generate-SummaryReport -AllSummaries $CompletedSummaries
        
        # Cleanup tracking files if everything completed
        if ($CompletedSummaries.Count -ge $TotalSearches) {
            Remove-Item $SearchTrackingFile -ErrorAction SilentlyContinue
            Remove-Item $CompletedTrackingFile -ErrorAction SilentlyContinue
            Log-Success "All $TotalSearches searches completed successfully!"
        }
        else {
            $remaining = $TotalSearches - $CompletedSummaries.Count
            Log-Warning "$remaining searches remaining. Run the script again to continue."
        }
    }
}

# Execute
Main
