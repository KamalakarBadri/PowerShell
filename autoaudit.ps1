

# Configuration
$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ClientSecret = "CyG8Q~FYt4sNxt5IejrMc2c24Ziz4a.t"

# Debug mode - SET TO $false FOR PRODUCTION!
$DebugMode = $True  # Shows full tokens when $true

# Calculate automatic date range (runs on 2nd of each month)
$CurrentDate = [DateTime]::UtcNow
if ($CurrentDate.Day -ne 2) {
    $CurrentDate = $CurrentDate.AddDays(2 - $CurrentDate.Day)  # For testing purposes
}

# Set start date to previous month 1st at 04:00:00 UTC
$StartDate = $CurrentDate.AddMonths(-1).AddDays(-$CurrentDate.AddMonths(-1).Day + 1).AddHours(4)

# Set end date to current month 1st at 03:59:59 UTC
$EndDate = $CurrentDate.AddDays(-$CurrentDate.Day + 1).AddHours(3).AddMinutes(59).AddSeconds(59)

# Format as ISO strings
$StartDate = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
$EndDate = $EndDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

# Sites and operations
$Sites = @(
    "https://geekbyteonline.sharepoint.com/sites/New365",
    "https://geekbyteonline.sharepoint.com/sites/2DayRetention",
    "https://geekbyteonline.sharepoint.com/sites/geekbyte",
    "https://geekbyteonline.sharepoint.com/sites/geetkteam",
    "https://geekbyteonline.sharepoint.com/sites/New365Site5"
)

$Operations = @("PageViewed", "FileAccessed", "FileDownloaded")
$ServiceFilter = "SharePoint"

# Constants
$MAX_RETRIES = 5
$RETRY_DELAY_SECONDS = 30      # Default wait time for rate limits
$THROTTLE_WAIT_SECONDS = 180   # 3 minutes wait time when 429 error occurs

# Logging function with colors
function Write-Log {
    param (
        [string]$Message,
        [string]$Color = "White"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $colorMapping = @{
        "White" = "White"
        "Red" = "Red"
        "Green" = "Green"
        "Yellow" = "Yellow"
        "Cyan" = "Cyan"
    }
    
    if ($colorMapping.ContainsKey($Color)) {
        Write-Host "[$timestamp] $Message" -ForegroundColor $colorMapping[$Color]
    } else {
        Write-Host "[$timestamp] $Message"
    }
}

# Display token information
function Show-TokenInfo {
    param (
        [string]$Token,
        [string]$TokenGenerationTime
    )
    
    if ($DebugMode) {
        Write-Log "=== SECURITY WARNING: FULL TOKEN DISPLAYED ===" -Color "Red"
        Write-Log "FULL TOKEN: $Token" -Color "Red"
        Write-Log "=== NEVER SHARE THIS TOKEN OR COMMIT TO VCS ===" -Color "Red"
    } else {
        $maskedToken = $Token.Substring(0, 10) + "..." + $Token.Substring($Token.Length - 10)
        Write-Log "Current token (masked): $maskedToken" -Color "Cyan"
    }
    
    if ($TokenGenerationTime) {
        $now = [DateTime]::UtcNow
        $tokenTime = [DateTime]::Parse($TokenGenerationTime)
        $age = $now - $tokenTime
        Write-Log "Token generated at: $TokenGenerationTime (age: $($age.TotalSeconds.ToString('0')) seconds)" -Color "Cyan"
    }
}

# 1. Get Access Token with retry logic
function Get-AccessToken {
    $AuthUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $Body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    $retryCount = 0
    while ($retryCount -lt $MAX_RETRIES) {
        try {
            $TokenResponse = Invoke-RestMethod -Uri $AuthUrl -Method Post -ContentType "application/x-www-form-urlencoded" -Body $Body
            Write-Log "New access token acquired" -Color "Green"
            return $TokenResponse.access_token, (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
        }
        catch {
            $retryCount++
            if ($retryCount -ge $MAX_RETRIES) {
                Write-Log "Failed to get access token after $MAX_RETRIES attempts: $_" -Color "Red"
                exit
            }
            
            Write-Log "Failed to get access token (attempt $retryCount/$MAX_RETRIES): $_" -Color "Yellow"
            Start-Sleep -Seconds $RETRY_DELAY_SECONDS
        }
    }
}

# 2. Create Audit Search with throttling
function New-AuditSearch {
    param (
        [string]$AccessToken,
        [string]$Site,
        [string]$Operation
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

    $retryCount = 0
    while ($retryCount -lt $MAX_RETRIES) {
        try {
            Write-Log "Creating search for $Site - $Operation" -Color "Cyan"
            $SearchQuery = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/security/auditLog/queries" `
                                            -Method Post `
                                            -Headers $Headers `
                                            -Body ($SearchParams | ConvertTo-Json -Depth 5)
            
            Write-Log "Search created successfully" -Color "Green"
            return $SearchQuery.id
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429) {
                Write-Log "Too Many Requests (429) - Waiting $THROTTLE_WAIT_SECONDS seconds before retrying..." -Color "Red"
                Start-Sleep -Seconds $THROTTLE_WAIT_SECONDS
                $retryCount++
                continue
            }
            elseif ($_.Exception.Response.StatusCode -eq 401) {
                Write-Log "Token expired, refreshing..." -Color "Yellow"
                $script:AccessToken, $script:TokenGenerationTime = Get-AccessToken
                $Headers.Authorization = "Bearer $script:AccessToken"
                continue
            }
            else {
                Write-Log "Failed to create audit search for $Site - $Operation : $_" -Color "Red"
                $retryCount++
                Start-Sleep -Seconds $RETRY_DELAY_SECONDS
                continue
            }
        }
    }
    
    Write-Log "Max retries ($MAX_RETRIES) exceeded for creating search" -Color "Red"
    return $null
}

# 3. Check Search Status with error handling
function Get-SearchStatus {
    param (
        [string]$AccessToken,
        [string]$SearchId
    )

    $Headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }

    $retryCount = 0
    while ($retryCount -lt $MAX_RETRIES) {
        try {
            $SearchStatus = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/security/auditLog/queries/$SearchId" `
                                             -Method Get `
                                             -Headers $Headers
            
            if ($SearchStatus.status -eq "succeeded" -or $SearchStatus.status -eq "failed") {
                Write-Log "Search $SearchId completed with status: $($SearchStatus.status)" -Color "Green"
            }

            return $SearchStatus.status
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429) {
                Write-Log "Too Many Requests (429) - Waiting $THROTTLE_WAIT_SECONDS seconds before retrying..." -Color "Red"
                Start-Sleep -Seconds $THROTTLE_WAIT_SECONDS
                $retryCount++
                continue
            }
            elseif ($_.Exception.Response.StatusCode -eq 401) {
                Write-Log "Token expired, refreshing..." -Color "Yellow"
                $script:AccessToken, $script:TokenGenerationTime = Get-AccessToken
                $Headers.Authorization = "Bearer $script:AccessToken"
                continue
            }
            else {
                Write-Log "Error checking status, will retry in 2 seconds: $_" -Color "Yellow"
                Start-Sleep -Seconds 2
                $retryCount++
                continue
            }
        }
    }
    
    Write-Log "Max retries ($MAX_RETRIES) exceeded for checking search status" -Color "Red"
    return "failed"
}

# 4. Retrieve Records with error handling
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
        $retryCount = 0
        $success = $false
        
        while (-not $success -and $retryCount -lt $MAX_RETRIES) {
            try {
                $Response = Invoke-RestMethod -Uri $Uri -Method Get -Headers $Headers
                $AllRecords += $Response.value
                $success = $true
                
                if ($Response.'@odata.nextLink') {
                    Write-Log "Retrieved $($AllRecords.Count) records..." -Color "Yellow"
                    $Uri = $Response.'@odata.nextLink'
                    Start-Sleep -Milliseconds 500
                }
                else {
                    $Uri = $null
                }
            }
            catch {
                if ($_.Exception.Response.StatusCode -eq 429) {
                    Write-Log "Too Many Requests (429) - Waiting $THROTTLE_WAIT_SECONDS seconds before retrying..." -Color "Red"
                    Start-Sleep -Seconds $THROTTLE_WAIT_SECONDS
                    $retryCount++
                    continue
                }
                elseif ($_.Exception.Response.StatusCode -eq 401) {
                    Write-Log "Token expired, refreshing..." -Color "Yellow"
                    $script:AccessToken, $script:TokenGenerationTime = Get-AccessToken
                    $Headers.Authorization = "Bearer $script:AccessToken"
                    $retryCount++
                    continue
                }
                else {
                    Write-Log "Error retrieving records, will retry in 2 seconds: $_" -Color "Yellow"
                    Start-Sleep -Seconds 2
                    $retryCount++
                    continue
                }
            }
        }
        
        if (-not $success) {
            Write-Log "Max retries ($MAX_RETRIES) exceeded for retrieving records" -Color "Red"
            break
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
        Write-Log "No records found for $Operation in $Site" -Color "Yellow"
        return 0
    }

    $SafeSite = ($Site -replace 'https?://|/', '_') -replace '[^a-zA-Z0-9_]', ''
    $StartDay = $StartDate.Split('T')[0].Replace('-', '')
    $EndDay = $EndDate.Split('T')[0].Replace('-', '')
    $CurrentTime = Get-Date -Format "yyyyMMdd_HHmmss"
    $FileName = "${SafeSite}_${StartDay}_${EndDay}_${Operation}_${CurrentTime}.csv"

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
        $Report | Export-Csv $FileName -NoTypeInformation -Force
        Write-Log "Saved $($Records.Count) records to $FileName" -Color "Green"
        return $Records.Count
    }
    catch {
        Write-Log "Failed to save CSV: $_" -Color "Red"
        return 0
    }
}

# Generate summary report
function Generate-SummaryReport {
    param (
        [array]$SummaryData
    )

    if ($SummaryData.Count -eq 0) {
        Write-Log "No data available for summary" -Color "Yellow"
        return
    }

    $CurrentTime = Get-Date -Format "yyyyMMdd_HHmmss"
    $SummaryFile = "AuditSummary_${CurrentTime}.csv"
    
    try {
        $SummaryData | Export-Csv $SummaryFile -NoTypeInformation -Force
        Write-Log "Summary file generated: $SummaryFile" -Color "Green"
    }
    catch {
        Write-Log "Failed to generate summary: $_" -Color "Red"
    }
}

# Load state from JSON file
function Load-State {
    param (
        [string]$FileName
    )

    if (Test-Path $FileName) {
        try {
            $content = Get-Content $FileName -Raw | ConvertFrom-Json -AsHashtable
            Write-Log "Loaded state from $FileName" -Color "Cyan"
            return $content
        }
        catch {
            Write-Log "Failed to load state file $FileName : $_" -Color "Yellow"
            return @{}
        }
    }
    return @{}
}

# Save state to JSON file
function Save-State {
    param (
        [string]$FileName,
        [hashtable]$Data
    )

    try {
        $Data | ConvertTo-Json -Depth 5 | Out-File $FileName -Force
        Write-Log "Saved state to $FileName" -Color "Cyan"
    }
    catch {
        Write-Log "Failed to save state file $FileName : $_" -Color "Red"
    }
}

# Archive search IDs file
function Archive-SearchIds {
    if (Test-Path "search_ids.json") {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $newName = "searchIds_${timestamp}.json"
        try {
            Rename-Item "search_ids.json" $newName
            Write-Log "Archived search_ids.json as $newName" -Color "Green"
            return $true
        }
        catch {
            Write-Log "Failed to archive search_ids.json : $_" -Color "Red"
            return $false
        }
    }
    return $true
}

# Main Execution
function Main {
    Write-Log "Starting audit log collection process..." -Color "Cyan"
    Write-Log "Date range: $StartDate to $EndDate" -Color "Cyan"
    
    # Authentication
    Write-Log "Authenticating..." -Color "Yellow"
    $script:AccessToken, $script:TokenGenerationTime = Get-AccessToken
    if (-not $script:AccessToken) {
        Write-Log "Failed to authenticate, exiting..." -Color "Red"
        exit
    }
    
    Show-TokenInfo -Token $script:AccessToken -TokenGenerationTime $script:TokenGenerationTime
    
    # Load existing state
    $existingSearches = Load-State "search_ids.json"
    $completedSearches = Load-State "completed_searches.json"
    Write-Log "Loaded $($existingSearches.Count) existing searches and $($completedSearches.Count) completed searches" -Color "Yellow"
    
    $summaryData = @()
    $totalSearches = $Sites.Count * $Operations.Count
    $createdSearches = 0
    
    # First create any missing searches
    foreach ($site in $Sites) {
        foreach ($operation in $Operations) {
            $key = "${site}_${operation}"
            
            if ($completedSearches.ContainsKey($key)) {
                $createdSearches++
                continue
            }
            
            if (-not $existingSearches.ContainsKey($key)) {
                $searchId = New-AuditSearch -AccessToken $script:AccessToken -Site $site -Operation $operation
                
                if ($searchId) {
                    $existingSearches[$key] = @{
                        SearchId = $searchId
                        CreatedTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    }
                    Save-State "search_ids.json" $existingSearches
                    $createdSearches++
                }
            }
            else {
                $createdSearches++
            }
            
            # Progress update
            $percentComplete = ($createdSearches / $totalSearches) * 100
            Write-Progress -Activity "Creating searches" -Status "Progress: $createdSearches of $totalSearches" -PercentComplete $percentComplete
        }
    }
    
    # Then process all searches (existing and new)
    $searchKeys = @($existingSearches.Keys | Where-Object { -not $completedSearches.ContainsKey($_) })
    $processedSearches = 0
    
    foreach ($key in $searchKeys) {
        $searchId = $existingSearches[$key].SearchId
        $status = $null
        
        do {
            $status = Get-SearchStatus -AccessToken $script:AccessToken -SearchId $searchId
            
            if ($status -eq "succeeded") {
                # Extract site and operation from key
                $parts = $key -split "_", 2
                $site = $parts[0]
                $operation = $parts[1]
                
                # Retrieve records
                Write-Log "Retrieving records for $site - $operation" -Color "Cyan"
                $records = Get-AuditRecords -AccessToken $script:AccessToken -SearchId $searchId
                
                # Save to CSV
                $recordCount = Save-AuditToCsv -Records $records -Site $site -Operation $operation
                $summaryData += @{
                    Site = $site.Split('/')[-1]
                    Operation = $operation
                    RecordCount = $recordCount
                }
                
                # Mark as completed
                $completedSearches[$key] = @{
                    CompletedTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    RecordCount = $recordCount
                }
                Save-State "completed_searches.json" $completedSearches
                
                $processedSearches++
                $percentComplete = ($processedSearches / $searchKeys.Count) * 100
                Write-Progress -Activity "Processing searches" -Status "Progress: $processedSearches of $($searchKeys.Count)" -PercentComplete $percentComplete
            }
            elseif ($status -eq "failed") {
                Write-Log "Search $searchId failed" -Color "Red"
                $completedSearches[$key] = @{
                    CompletedTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    Status = "failed"
                }
                Save-State "completed_searches.json" $completedSearches
                $processedSearches++
            }
            else {
                # Still running, wait before checking again
                Write-Log "Search $searchId status: $status. Waiting $RETRY_DELAY_SECONDS seconds..." -Color "Yellow"
                Start-Sleep -Seconds $RETRY_DELAY_SECONDS
            }
        } while ($status -ne "succeeded" -and $status -ne "failed")
    }
    
    # Generate summary report
    Generate-SummaryReport -SummaryData $summaryData
    
    # Archive search IDs if all completed
    if ($completedSearches.Count -eq $totalSearches) {
        if (Archive-SearchIds) {
            # Only remove if archiving was successful
            if (Test-Path "completed_searches.json") {
                try {
                    Remove-Item "completed_searches.json"
                }
                catch {
                    Write-Log "Failed to remove completed_searches.json: $_" -Color "Yellow"
                }
            }
        }
        
        Write-Log "All operations completed successfully!" -Color "Green"
    }
    else {
        $remaining = $totalSearches - $completedSearches.Count
        Write-Log "$remaining searches remaining. Run the script again to continue." -Color "Yellow"
    }
}

# Execute
Main
