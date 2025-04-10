<#
.SYNOPSIS
    SharePoint Permission Scanner with Logging and Progress Tracking
.DESCRIPTION
    Scans SharePoint document libraries for item-level permissions, generates CSV reports,
    and provides detailed logging with progress visualization.
.NOTES
    Version: 3.0
    Author: Your Name
    Date: $(Get-Date -Format "yyyy-MM-dd")
#>

param(
    [string]$LogPath = ".\PermissionScanLog_$(Get-Date -Format 'yyyyMMdd-HHmmss').log",
    [switch]$SkipItemCount
)

#region Initialization
$scriptStartTime = Get-Date
$ErrorActionPreference = "Stop"

# Initialize logging
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogPath -Value $logEntry
    
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        "DEBUG" { "Gray" }
        default { "White" }
    }
    Write-Host $logEntry -ForegroundColor $color
}

# Create log file
New-Item -Path $LogPath -ItemType File -Force | Out-Null
Write-Log "Starting SharePoint Permission Scanner"
Write-Log "Log file created at $LogPath"
#endregion

#region Configuration
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
#endregion

#region Main Processing
foreach ($siteUrl in $SiteUrls) {
    $siteStartTime = Get-Date
    Write-Log "`n===================================================================="
    Write-Log "PROCESSING SITE: $siteUrl"
    Write-Log "===================================================================="
    
    # Connect to SharePoint
    try {
        Write-Log "Connecting to SharePoint site..."
        Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId -ErrorAction Stop
        Write-Log "Successfully connected to $siteUrl"
    }
    catch {
        Write-Log "Failed to connect to $siteUrl : $_" -Level "ERROR"
        continue
    }

    # Get all document libraries
    try {
        Write-Log "Retrieving document libraries..."
        $lists = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists" -Method Get -ErrorAction Stop
        Write-Log "Found $($lists.value.Count) lists in the site"
    }
    catch {
        Write-Log "Error retrieving lists: $_" -Level "ERROR"
        continue
    }

    $reportData = @()
    $processedItems = 0
    $itemsWithPermissions = 0
    $totalItems = 0

    # Count total items if not skipped
    if (-not $SkipItemCount) {
        Write-Log "Counting items in all document libraries..."
        foreach ($list in $lists.value) {
            if ($list.BaseTemplate -ne 101 -or $ExcludedLists -contains $list.Title) { continue }
            
            try {
                $listItems = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/ItemCount" -Method Get
                $totalItems += $listItems.value
                Write-Log "  List $($list.Title) has $($listItems.value) items" -Level "DEBUG"
            }
            catch {
                Write-Log "Error counting items in list $($list.Title): $_" -Level "WARN"
            }
        }
        Write-Log "Total items to process: $totalItems"
    }
    else {
        Write-Log "Skipping item count (using approximate progress)" -Level "WARN"
        $totalItems = 1000 # Default for progress tracking
    }

    # Initialize progress tracking
    $progressParams = @{
        Activity = "Scanning $($siteUrl.Split('/')[-1])"
        Status = "Preparing to scan"
        PercentComplete = 0
    }

    # Process each list
    foreach ($list in $lists.value) {
        if ($list.BaseTemplate -ne 101) { 
            Write-Log "Skipping non-document library: $($list.Title) (BaseTemplate: $($list.BaseTemplate))" -Level "DEBUG"
            continue 
        }

        if ($ExcludedLists -contains $list.Title) {
            Write-Log "Skipping excluded list: $($list.Title)" -Level "DEBUG"
            continue
        }

        Write-Log "Processing Document Library: $($list.Title)"

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
                    
                    # Update progress (handle cases where we skipped item count)
                    if ($totalItems -gt 0) {
                        $percentComplete = [math]::Min(100, [math]::Round(($processedItems / $totalItems) * 100, 2))
                    }
                    else {
                        $percentComplete = [math]::Min(100, ($processedItems % 100))
                    }
                    
                    $progressParams.Status = "Item $processedItems"
                    $progressParams.PercentComplete = $percentComplete
                    $progressParams.CurrentOperation = "$($list.Title) - $($item.Id)"
                    Write-Progress @progressParams

                    try {
                        # Get item details and determine type
                        $fileSystemObjectType = $item.FileSystemObjectType
                        $itemName = $null
                        $itemLocation = $null
                        $itemSize = $null
                        
                        if ($fileSystemObjectType -eq 0) {
                            $itemType = "File"
                            try {
                                $fileResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/file" -Method Get -ErrorAction Stop
                                $itemName = $fileResponse.Name
                                $itemLocation = $fileResponse.ServerRelativeUrl
                                $itemSize = $fileResponse.Length
                                Write-Log "    Processing file: $itemName" -Level "DEBUG"
                            }
                            catch {
                                Write-Log "    Error getting file details for item $($item.Id): $_" -Level "ERROR"
                                continue
                            }
                        } 
                        elseif ($fileSystemObjectType -eq 1) {
                            $itemType = "Folder"
                            try {
                                $folderResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/folder" -Method Get -ErrorAction Stop
                                $itemName = $folderResponse.Name
                                $itemLocation = $folderResponse.ServerRelativeUrl
                                Write-Log "    Processing folder: $itemName" -Level "DEBUG"
                            }
                            catch {
                                Write-Log "    Error getting folder details for item $($item.Id): $_" -Level "ERROR"
                                continue
                            }
                        } 
                        else {
                            $itemType = "ListItem"
                            $itemName = $item.Title
                            $itemLocation = $null
                            Write-Log "    Processing list item: $itemName" -Level "DEBUG"
                        }

                        # Check if item has unique permissions
                        try {
                            $uniquePerms = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/HasUniqueRoleAssignments" -Method Get -ErrorAction Stop
                            $hasUniquePerms = $uniquePerms.value
                            Write-Log "    Unique permissions: $hasUniquePerms" -Level "DEBUG"
                        }
                        catch {
                            Write-Log "    Error checking unique permissions for item $($item.Id): $_" -Level "ERROR"
                            $hasUniquePerms = $false
                        }

                        # Get permissions info
                        try {
                            $permsInfo = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation" -Method Get -ErrorAction Stop
                            Write-Log "    Retrieved permissions information" -Level "DEBUG"
                        }
                        catch {
                            Write-Log "    Error getting permissions info for item $($item.Id): $_" -Level "ERROR"
                            continue
                        }

                        # Initialize permission collections
                        $readUsers = @()
                        $editUsers = @()
                        $fullControlUsers = @()
                        $sharingLinks = @()

                        # Process direct permissions and group memberships
                        if ($permsInfo.permissionsInformation.principals) {
                            Write-Log "    Found $($permsInfo.permissionsInformation.principals.Count) principals" -Level "DEBUG"
                            
                            foreach ($principalElement in $permsInfo.permissionsInformation.principals) {
                                $principal = $principalElement.principal
                                $role = $principalElement.role
                                Write-Log "      Processing principal: $($principal.name) (Type: $($principal.principalType)) with role: $role" -Level "DEBUG"

                                if ($principal.principalType -eq 1) {
                                    # User - add to appropriate permission collection
                                    $principalUpn = $principal.userPrincipalName ?? $principal.email
                                    $principalName = $principal.name ?? $principalUpn
                                    Write-Log "        User principal: $principalName" -Level "DEBUG"
                                    
                                    switch ($role) {
                                        1 { $readUsers += $principalName }
                                        2 { $editUsers += $principalName }
                                        3 { $fullControlUsers += $principalName }
                                    }
                                }
                                elseif ($principal.principalType -in @(4,8)) {
                                    # Group - get ALL members and add them with group name
                                    $groupName = $principal.name
                                    Write-Log "      Processing group: $groupName (Type: $($principal.principalType))" -Level "DEBUG"
                                    
                                    $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))/Users"
                                    try {
                                        Write-Log "        Retrieving members for group $groupName..." -Level "DEBUG"
                                        $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                                        Write-Log "        Found $($members.value.Count) members in group $groupName" -Level "DEBUG"
                                        
                                        if ($members.value.Count -eq 0) {
                                            Write-Log "        No members found in group $groupName - adding group name only" -Level "WARN"
                                            switch ($role) {
                                                1 { $readUsers += "$groupName [no members found]" }
                                                2 { $editUsers += "$groupName [no members found]" }
                                                3 { $fullControlUsers += "$groupName [no members found]" }
                                            }
                                            continue
                                        }

                                        foreach ($member in $members.value) {
                                            Write-Log "          Processing member (Type: $($member.PrincipalType))..." -Level "DEBUG"
                                            $memberName = $null
                                            
                                            if ($member.PrincipalType -eq 1) {
                                                # User
                                                $memberName = $member.Title ?? ($member.UserPrincipalName ?? $member.Email)
                                                Write-Log "            User member: $memberName" -Level "DEBUG"
                                            }
                                            elseif ($member.PrincipalType -eq 4) {
                                                # Security Group
                                                $memberName = $member.Title ?? $member.LoginName
                                                Write-Log "            Security Group member: $memberName" -Level "DEBUG"
                                            }
                                            elseif ($member.PrincipalType -eq 8) {
                                                # SharePoint Group
                                                $memberName = $member.Title ?? $member.LoginName
                                                Write-Log "            SharePoint Group member: $memberName" -Level "DEBUG"
                                            }
                                            else {
                                                # Other principal types
                                                $memberName = $member.Title ?? $member.LoginName ?? $member.Email ?? "Unknown Principal"
                                                Write-Log "            Other member type ($($member.PrincipalType)): $memberName" -Level "DEBUG"
                                            }
                                            
                                            # If member name is still empty, get the group title
                                            if ([string]::IsNullOrEmpty($memberName)) {
                                                Write-Log "            Member name empty, trying to get group title..." -Level "WARN"
                                                try {
                                                    $groupInfo = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))" -Method Get -ErrorAction Stop
                                                    $memberName = $groupInfo.Title ?? $groupInfo.LoginName ?? $groupName
                                                    Write-Log "            Using group title: $memberName" -Level "DEBUG"
                                                }
                                                catch {
                                                    $memberName = $groupName
                                                    Write-Log "            Using original group name: $memberName" -Level "DEBUG"
                                                }
                                            }
                                            
                                            if ($memberName) {
                                                switch ($role) {
                                                    1 { $readUsers += "$memberName (via $groupName)" }
                                                    2 { $editUsers += "$memberName (via $groupName)" }
                                                    3 { $fullControlUsers += "$memberName (via $groupName)" }
                                                }
                                            }
                                        }
                                    }
                                    catch {
                                        Write-Log "Error getting members for group $groupName : $_" -Level "ERROR"
                                        switch ($role) {
                                            1 { $readUsers += "$groupName [members not accessible]" }
                                            2 { $editUsers += "$groupName [members not accessible]" }
                                            3 { $fullControlUsers += "$groupName [members not accessible]" }
                                        }
                                    }
                                }
                            }
                        }

                        # Process sharing links
                        if ($permsInfo.permissionsInformation.links) {
                            Write-Log "    Found $($permsInfo.permissionsInformation.links.Count) sharing links" -Level "DEBUG"
                            
                            foreach ($link in $permsInfo.permissionsInformation.links) {
                                if ($link.linkDetails -and $link.linkDetails.Url) {
                                    $linkUrl = $link.linkDetails.Url
                                    $linkType = if ($link.linkDetails.IsEditLink -or $link.linkDetails.IsReviewLink) { "Edit" } else { "Read" }
                                    $sharingLinks += "$linkUrl ($linkType access)"
                                    Write-Log "      Sharing link: $linkUrl ($linkType access)" -Level "DEBUG"
                                    
                                    if ($link.linkMembers) {
                                        Write-Log "        Found $($link.linkMembers.Count) link members" -Level "DEBUG"
                                        foreach ($member in $link.linkMembers) {
                                            $memberUpn = $member.userPrincipalName ?? $member.email
                                            $memberName = $member.displayName ?? $memberUpn
                                            Write-Log "          Link member: $memberName" -Level "DEBUG"
                                            
                                            if ($linkType -eq "Edit") {
                                                $editUsers += "$memberName (via sharing link : $linkUrl)"
                                            } else {
                                                $readUsers += "$memberName (via sharing link : $linkUrl)"
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        # Add to report if there are any permissions
                        if ($readUsers.Count -gt 0 -or $editUsers.Count -gt 0 -or $fullControlUsers.Count -gt 0 -or $sharingLinks.Count -gt 0) {
                            $itemsWithPermissions++
                            $reportEntry = [PSCustomObject]@{
                                SiteName         = $siteUrl.Split("/")[-1]
                                LibraryName      = $list.Title
                                ItemID           = $item.Id
                                ItemType         = $itemType
                                Name             = $itemName
                                Location         = $itemLocation
                                Size             = if ($itemSize) { "$([math]::Round($itemSize/1KB, 2)) KB" } else { "NA" }
                                ReadUsers        = ($readUsers | Sort-Object -Unique) -join "`n"
                                EditUsers        = ($editUsers | Sort-Object -Unique) -join "`n"
                                FullControlUsers = ($fullControlUsers | Sort-Object -Unique) -join "`n"
                                SharingLinks     = if ($sharingLinks.Count -gt 0) { ($sharingLinks | Sort-Object -Unique) -join "`n" } else { "" }
                                UniquePerms      = if ($hasUniquePerms) { "Yes" } else { "No" }
                                LastModified     = if ($item.Modified) { $item.Modified } else { "" }
                            }
                            $reportData += $reportEntry
                            
                            Write-Log "      Found permissions for $itemType $itemName" -Level "INFO"
                            Write-Log "        Read Users: $($readUsers.Count)" -Level "DEBUG"
                            Write-Log "        Edit Users: $($editUsers.Count)" -Level "DEBUG"
                            Write-Log "        Full Control Users: $($fullControlUsers.Count)" -Level "DEBUG"
                            Write-Log "        Sharing Links: $($sharingLinks.Count)" -Level "DEBUG"
                        }
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
            Write-Log "REPORT SUMMARY:"
            Write-Log "  Total items processed: $processedItems"
            Write-Log "  Items with permissions recorded: $itemsWithPermissions"
            Write-Log "  Report generated: $fileName"
        }
        catch {
            Write-Log "Error generating report: $_" -Level "ERROR"
        }
    } else {
        Write-Log "No permissions found for items in $siteUrl."
        Write-Log "Total items processed: $processedItems"
    }

    $siteDuration = (Get-Date) - $siteStartTime
    Write-Log "Site processing completed in $($siteDuration.ToString('hh\:mm\:ss'))"
    
    Disconnect-PnPOnline
    Write-Log "Disconnected from $siteUrl"
    Write-Progress -Activity "Completed $($siteUrl.Split('/')[-1])" -Completed
}

$totalDuration = (Get-Date) - $scriptStartTime
Write-Log "`nScript completed processing all sites in $($totalDuration.ToString('hh\:mm\:ss'))"
Write-Log "Log file saved to $LogPath"
#endregion
