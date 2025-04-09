<#
.SYNOPSIS
    SharePoint Permission Reporter with Logging and Progress Tracking
.DESCRIPTION
    This script scans SharePoint document libraries and reports on item-level permissions,
    with detailed logging and progress visualization.
.NOTES
    Version: 2.0
    Author: Your Name
    Date: $(Get-Date -Format "yyyy-MM-dd")
#>

param(
    [string]$LogPath = ".\PermissionScanLog_$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
)

# Initialize logging
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogPath -Value $logEntry
    Write-Host $logEntry -ForegroundColor $(switch ($Level) { "ERROR" { "Red" } "WARN" { "Yellow" } default { "White" } })
}

# Configuration
$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ThumbPrint = "B799789F78628CAE56B4D0F380FD551EB754E0DB"

$SiteUrls = @(
    "https://geekbyteonline.sharepoint.com/sites/New365",
    "https://geekbyteonline.sharepoint.com/sites/AnotherSite",
    "https://geekbyteonline.sharepoint.com/sites/ThirdSite"
)

$ExcludedLists = @("Access Requests", "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", 
    "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms",
    "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", 
    "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Images", "site collection images",
    "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", 
    "Reusable Content", "Reporting Metadata", "Reporting Templates", "Search Config List", "Site Assets", 
    "Preservation Hold Library", "Solution Gallery", "Style Library", "Suggested Content Browser Locations", 
    "Theme Gallery", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", 
    "Workflow History", "Workflow Tasks", "Pages")

# Create log file
New-Item -Path $LogPath -ItemType File -Force | Out-Null
Write-Log "Starting SharePoint Permission Scanner"
Write-Log "Log file created at $LogPath"

# Main processing loop
foreach ($siteUrl in $SiteUrls) {
    $siteStartTime = Get-Date
    Write-Log "`n===================================================================="
    Write-Log "PROCESSING SITE: $siteUrl"
    Write-Log "===================================================================="
    
    try {
        Write-Log "Connecting to SharePoint site..."
        Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId -ErrorAction Stop
        Write-Log "Successfully connected to $siteUrl" -Level "INFO"
    }
    catch {
        Write-Log "Failed to connect to $siteUrl : $_" -Level "ERROR"
        continue
    }

    # Get all document libraries
    try {
        Write-Log "Retrieving document libraries..."
        $lists = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists" -Method Get -ErrorAction Stop
        Write-Log "Found $($lists.value.Count) lists in the site" -Level "INFO"
    }
    catch {
        Write-Log "Error retrieving lists: $_" -Level "ERROR"
        continue
    }

    $reportData = @()
    $processedItems = 0
    $itemsWithPermissions = 0
    $totalItems = 0

    # First pass to count total items
    Write-Log "Counting items in all document libraries..."
    foreach ($list in $lists.value) {
        if ($list.BaseTemplate -ne 101 -or $ExcludedLists -contains $list.Title) { continue }
        
        try {
            $listItems = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/ItemCount" -Method Get
            $totalItems += $listItems.value
        }
        catch {
            Write-Log "Error counting items in list $($list.Title): $_" -Level "WARN"
        }
    }
    Write-Log "Total items to process: $totalItems" -Level "INFO"

    # Initialize progress tracking
    $progressParams = @{
        Activity = "Scanning $siteUrl"
        Status = "Processing items"
        PercentComplete = 0
    }

    # Process each list
    foreach ($list in $lists.value) {
        if ($list.BaseTemplate -ne 101) { 
            Write-Log "Skipping non-document library: $($list.Title) (BaseTemplate: $($list.BaseTemplate))" -Level "INFO"
            continue 
        }

        if ($ExcludedLists -contains $list.Title) {
            Write-Log "Skipping excluded list: $($list.Title)" -Level "INFO"
            continue
        }

        Write-Log "`nProcessing Document Library: $($list.Title)" -Level "INFO"

        $nextPageUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items?`$top=500"
        $pageCount = 0
        
        do {
            $pageCount++
            Write-Log "  Retrieving page $pageCount of items..." -Level "DEBUG"
            
            try {
                $response = Invoke-PnPSPRestMethod -Url $nextPageUrl -Method Get -ErrorAction Stop
                $listItems = $response.value
                $nextPageUrl = $response."odata.nextLink"

                foreach ($item in $listItems) {
                    $processedItems++
                    $percentComplete = [math]::Min(100, [math]::Round(($processedItems / $totalItems) * 100, 2))
                    
                    # Update progress bar
                    $progressParams.Status = "Processing item $processedItems of $totalItems ($percentComplete%)"
                    $progressParams.PercentComplete = $percentComplete
                    $progressParams.CurrentOperation = "$($list.Title) - $($item.Title ?? $item.Id)"
                    Write-Progress @progressParams

                    try {
                        # [Previous permission checking logic goes here]
                        # ... (Include all the permission checking code from previous version)
                        # [End of permission checking logic]

                    } catch {
                        Write-Log "Error processing item $($item.Id): $_" -Level "ERROR"
                    }
                }
            } catch {
                Write-Log "Error retrieving items: $_" -Level "ERROR"
                $nextPageUrl = $null
            }
        } while ($nextPageUrl)
    }

    # Generate CSV report
    if ($reportData.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $siteName = $siteUrl.Split("/")[-1]
        $fileName = "$siteName-PermissionsReport_$timestamp.csv"
        
        try {
            $reportData | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            Write-Log "REPORT SUMMARY:" -Level "INFO"
            Write-Log "Total items processed: $processedItems" -Level "INFO"
            Write-Log "Items with permissions recorded: $itemsWithPermissions" -Level "INFO"
            Write-Log "Report generated: $fileName" -Level "INFO"
        }
        catch {
            Write-Log "Error generating report: $_" -Level "ERROR"
        }
    } else {
        Write-Log "No permissions found for items in $siteUrl." -Level "INFO"
        Write-Log "Total items processed: $processedItems" -Level "INFO"
    }

    $siteDuration = (Get-Date) - $siteStartTime
    Write-Log "Site processing completed in $($siteDuration.ToString('hh\:mm\:ss'))" -Level "INFO"
    
    Disconnect-PnPOnline
    Write-Log "Disconnected from $siteUrl" -Level "INFO"
    Write-Progress -Activity "Completed $siteUrl" -Completed
}

$totalDuration = (Get-Date) - $scriptStartTime
Write-Log "`nScript completed processing all sites in $($totalDuration.ToString('hh\:mm\:ss'))" -Level "INFO"
Write-Log "Log file saved to $LogPath" -Level "INFO"
