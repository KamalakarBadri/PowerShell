<#
.SYNOPSIS
    SharePoint Permission Auditor with Detailed Logging and Progress Tracking
.DESCRIPTION
    This script audits SharePoint document libraries for specific users' permissions,
    including direct permissions, group memberships, and sharing links.
    Features detailed logging and progress visualization.
.NOTES
    File Name      : SharePointPermissionAuditor.ps1
    Author         : GeekByte
    Prerequisite   : PnP.PowerShell module
#>

# Parameters
$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ThumbPrint = "B799789F78628CAE56B4D0F380FD551EB754E0DB"

# Array of site URLs to process
$SiteUrls = @(
    "https://geekbyteonline.sharepoint.com/sites/New365",
    "https://geekbyteonline.sharepoint.com/sites/AnotherSite",
    "https://geekbyteonline.sharepoint.com/sites/ThirdSite"
)

# Array of users to check (case-insensitive)
$UserUPNs = @(
    "nodownload@geekbyte.online",
    "read@geekbyte.online",
    "sharelink@geekbyte.online"
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

# Initialize global counters
$global:TotalItemsProcessed = 0
$global:TotalItemsWithAccess = 0
$global:StartTime = Get-Date

# Logging function
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [string]$ForegroundColor = "White"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to console
    Write-Host $logEntry -ForegroundColor $ForegroundColor
    
    # Optionally write to log file
    # $logEntry | Out-File -FilePath "PermissionAuditLog.txt" -Append
}

# Progress tracking function
function Show-Progress {
    param (
        [int]$Current,
        [int]$Total,
        [string]$Activity,
        [string]$Status
    )
    
    $percentComplete = ($Current / $Total) * 100
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $percentComplete
}

# Main script execution
Write-Log "SCRIPT STARTED" -Level "INFO" -ForegroundColor "Cyan"
Write-Log "Processing $($SiteUrls.Count) sites and $($UserUPNs.Count) users" -Level "INFO" -ForegroundColor "Cyan"

foreach ($siteUrl in $SiteUrls) {
    $siteStartTime = Get-Date
    Write-Log "`n================================================================" -Level "HEADER" -ForegroundColor "Cyan"
    Write-Log "PROCESSING SITE: $siteUrl" -Level "HEADER" -ForegroundColor "Cyan"
    Write-Log "================================================================" -Level "HEADER" -ForegroundColor "Cyan"
    
    try {
        Write-Log "Connecting to SharePoint site..." -Level "INFO" -ForegroundColor "Yellow"
        Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId -ErrorAction Stop
        Write-Log "Successfully connected to $siteUrl" -Level "SUCCESS" -ForegroundColor "Green"
    }
    catch {
        Write-Log "Failed to connect to $siteUrl : $_" -Level "ERROR" -ForegroundColor "Red"
        continue
    }

    # Get all document libraries
    Write-Log "Retrieving document libraries..." -Level "INFO" -ForegroundColor "Yellow"
    try {
        $lists = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists" -Method Get -ErrorAction Stop
        Write-Log "Found $($lists.value.Count) lists in the site" -Level "INFO" -ForegroundColor "Green"
    }
    catch {
        Write-Log "Error retrieving lists: $_" -Level "ERROR" -ForegroundColor "Red"
        continue
    }

    $reportData = @()
    $siteItemsProcessed = 0
    $siteItemsWithAccess = 0
    $totalItemsToProcess = 0

    # First pass to estimate total items (for progress tracking)
    Write-Log "Estimating total items to process..." -Level "INFO" -ForegroundColor "DarkGray"
    foreach ($list in $lists.value) {
        if ($list.BaseTemplate -ne 101) { continue }
        if ($ExcludedLists -contains $list.Title) { continue }
        
        try {
            $listItemCount = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/ItemCount" -Method Get -ErrorAction Stop
            $totalItemsToProcess += $listItemCount
        }
        catch {
            Write-Log "Could not get item count for list $($list.Title), using default estimate" -Level "WARNING" -ForegroundColor "Yellow"
            $totalItemsToProcess += 1000 # Default estimate if we can't get count
        }
    }

    Write-Log "Estimated total items to process in site: $totalItemsToProcess" -Level "INFO" -ForegroundColor "Green"

    foreach ($list in $lists.value) {
        if ($list.BaseTemplate -ne 101) { 
            Write-Log "Skipping non-document library: $($list.Title) (BaseTemplate: $($list.BaseTemplate))" -Level "DEBUG" -ForegroundColor "Gray"
            continue 
        }

        if ($ExcludedLists -contains $list.Title) {
            Write-Log "Skipping excluded list: $($list.Title)" -Level "DEBUG" -ForegroundColor "Gray"
            continue
        }

        $listStartTime = Get-Date
        Write-Log "`nProcessing Document Library: $($list.Title)" -Level "INFO" -ForegroundColor "Yellow"

        $nextPageUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items?`$top=1000"
        $pageCount = 0
        
        do {
            $pageCount++
            Write-Log "  Retrieving page $pageCount of items..." -Level "DEBUG" -ForegroundColor "DarkGray"
            
            try {
                $response = Invoke-PnPSPRestMethod -Url $nextPageUrl -Method Get -ErrorAction Stop
                $listItems = $response.value
                $nextPageUrl = $response."odata.nextLink"
                Write-Log "  Found $($listItems.Count) items on this page" -Level "DEBUG" -ForegroundColor "DarkGray"

                # Show progress for current page
                $activity = "Processing $($list.Title) (Page $pageCount)"
                $status = "Items processed: $siteItemsProcessed | With access: $siteItemsWithAccess"
                Show-Progress -Current $siteItemsProcessed -Total $totalItemsToProcess -Activity $activity -Status $status

                foreach ($item in $listItems) {
                    $siteItemsProcessed++
                    $global:TotalItemsProcessed++
                    
                    # Update progress every 10 items
                    if ($siteItemsProcessed % 10 -eq 0) {
                        $status = "Items processed: $siteItemsProcessed/$totalItemsToProcess | With access: $siteItemsWithAccess"
                        Show-Progress -Current $siteItemsProcessed -Total $totalItemsToProcess -Activity $activity -Status $status
                    }

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
                                Write-Log "    Processing file: $itemName" -Level "DEBUG" -ForegroundColor "DarkGray"
                            }
                            catch {
                                Write-Log "    Error getting file details for item $($item.Id): $_" -Level "ERROR" -ForegroundColor "Red"
                                continue
                            }
                        } 
                        elseif ($fileSystemObjectType -eq 1) {
                            $itemType = "Folder"
                            try {
                                $folderResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/folder" -Method Get -ErrorAction Stop
                                $itemName = $folderResponse.Name
                                $itemLocation = $folderResponse.ServerRelativeUrl
                                Write-Log "    Processing folder: $itemName" -Level "DEBUG" -ForegroundColor "DarkGray"
                            }
                            catch {
                                Write-Log "    Error getting folder details for item $($item.Id): $_" -Level "ERROR" -ForegroundColor "Red"
                                continue
                            }
                        } 
                        else {
                            $itemType = "ListItem"
                            $itemName = $item.Title
                            $itemLocation = $null
                            Write-Log "    Processing list item: $itemName" -Level "DEBUG" -ForegroundColor "DarkGray"
                        }

                        # Check if item has unique permissions
                        try {
                            $uniquePerms = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/HasUniqueRoleAssignments" -Method Get -ErrorAction Stop
                            $hasUniquePerms = $uniquePerms.value
                        }
                        catch {
                            Write-Log "    Error checking unique permissions for item $($item.Id): $_" -Level "WARNING" -ForegroundColor "Yellow"
                            $hasUniquePerms = $false
                        }

                        # Get permissions info (both direct and inherited)
                        try {
                            $permsInfo = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation" -Method Get -ErrorAction Stop
                        }
                        catch {
                            Write-Log "    Error getting permissions info for item $($item.Id): $_" -Level "ERROR" -ForegroundColor "Red"
                            continue
                        }

                        foreach ($UserUPN in $UserUPNs) {
                            $readSources = @()
                            $editSources = @()
                            $fullControl = $false
                            $sharingLinks = @()

                            # Process direct permissions and group memberships
                            if ($permsInfo.permissionsInformation.principals) {
                                Write-Log "    Found $($permsInfo.permissionsInformation.principals.Count) principals with permissions" -Level "DEBUG" -ForegroundColor "DarkGray"
                                
                                foreach ($principalElement in $permsInfo.permissionsInformation.principals) {
                                    $principal = $principalElement.principal
                                    $role = $principalElement.role
                                    Write-Log "      Processing principal: $($principal.name) (Type: $($principal.principalType)) with role: $role" -Level "DEBUG" -ForegroundColor "DarkGray"

                                    if ($principal.principalType -eq 1) {
                                        # User - check if it matches our target UPN
                                        $principalUpn = $principal.userPrincipalName ?? $principal.email
                                        $principalName = $principal.name ?? $principalUpn
                                        Write-Log "        User principal found: $principalName" -Level "DEBUG" -ForegroundColor "DarkGray"
                                        
                                        if ($principalUpn -like "*$UserUPN*") {
                                            switch ($role) {
                                                1 { $readSources += "Direct Permission: $principalName" }
                                                2 { $editSources += "Direct Permission: $principalName" }
                                                3 { $fullControl = $true }
                                            }
                                        }
                                    }
                                    elseif ($principal.principalType -in @(4,8)) {
                                        # Group - get ALL members and check if any match our target UPN
                                        $groupName = $principal.name
                                        Write-Log "      Processing group: $groupName (Type: $($principal.principalType))" -Level "DEBUG" -ForegroundColor "DarkGray"
                                        
                                        $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))/Users"
                                        try {
                                            Write-Log "        Retrieving members for group $groupName..." -Level "DEBUG" -ForegroundColor "DarkGray"
                                            $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                                            Write-Log "        Found $($members.value.Count) members in group $groupName" -Level "DEBUG" -ForegroundColor "DarkGray"
                                            
                                            if ($members.value.Count -eq 0) {
                                                Write-Log "        No members found in group $groupName" -Level "WARNING" -ForegroundColor "Yellow"
                                                continue
                                            }

                                            foreach ($member in $members.value) {
                                                Write-Log "          Processing member (Type: $($member.PrincipalType))..." -Level "DEBUG" -ForegroundColor "DarkGray"
                                                $memberName = $null
                                                $memberUpn = $null
                                                
                                                if ($member.PrincipalType -eq 1) {
                                                    # User
                                                    $memberUpn = $member.UserPrincipalName ?? $member.Email
                                                    $memberName = $member.Title ?? $memberUpn
                                                    Write-Log "            User member: $memberName" -Level "DEBUG" -ForegroundColor "DarkGray"
                                                }
                                                elseif ($member.PrincipalType -eq 4) {
                                                    # Security Group
                                                    $memberName = $member.Title ?? $member.LoginName
                                                    Write-Log "            Security Group member: $memberName" -Level "DEBUG" -ForegroundColor "DarkGray"
                                                }
                                                elseif ($member.PrincipalType -eq 8) {
                                                    # SharePoint Group
                                                    $memberName = $member.Title ?? $member.LoginName
                                                    Write-Log "            SharePoint Group member: $memberName" -Level "DEBUG" -ForegroundColor "DarkGray"
                                                }
                                                else {
                                                    # Other principal types
                                                    $memberName = $member.Title ?? $member.LoginName ?? $member.Email ?? "Unknown Principal"
                                                    Write-Log "            Other member type ($($member.PrincipalType)): $memberName" -Level "DEBUG" -ForegroundColor "DarkGray"
                                                }
                                                
                                                if ($member.PrincipalType -eq 1 -and $memberUpn -like "*$UserUPN*") {
                                                    switch ($role) {
                                                        1 { $readSources += "Group Membership ($groupName): $memberName" }
                                                        2 { $editSources += "Group Membership ($groupName): $memberName" }
                                                        3 { $fullControl = $true }
                                                    }
                                                }
                                            }
                                        }
                                        catch {
                                            Write-Log "        Error getting members for group $groupName : $_" -Level "WARNING" -ForegroundColor "Yellow"
                                        }
                                    }
                                }
                            }

                            # Process sharing links
                            if ($permsInfo.permissionsInformation.links) {
                                Write-Log "    Found $($permsInfo.permissionsInformation.links.Count) sharing links" -Level "DEBUG" -ForegroundColor "DarkGray"
                                
                                foreach ($link in $permsInfo.permissionsInformation.links) {
                                    if ($link.linkDetails -and $link.linkDetails.Url) {
                                        $linkUrl = $link.linkDetails.Url
                                        $linkType = if ($link.linkDetails.IsEditLink -or $link.linkDetails.IsReviewLink) { "Edit" } else { "Read" }
                                        $sharingLinks += "$linkUrl ($linkType access)"
                                        Write-Log "      Found sharing link: $linkUrl ($linkType access)" -Level "DEBUG" -ForegroundColor "DarkGray"
                                        
                                        if ($link.linkMembers) {
                                            Write-Log "        Found $($link.linkMembers.Count) link members" -Level "DEBUG" -ForegroundColor "DarkGray"
                                            foreach ($member in $link.linkMembers) {
                                                $memberUpn = $member.userPrincipalName ?? $member.email
                                                $memberName = $member.displayName ?? $memberUpn
                                                Write-Log "          Link member: $memberName" -Level "DEBUG" -ForegroundColor "DarkGray"
                                                
                                                if ($memberUpn -like "*$UserUPN*") {
                                                    if ($linkType -eq "Edit") {
                                                        $editSources += "Sharing Link ($linkUrl): $memberName"
                                                    } else {
                                                        $readSources += "Sharing Link ($linkUrl): $memberName"
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if ($readSources -or $editSources -or $fullControl) {
                                $siteItemsWithAccess++
                                $global:TotalItemsWithAccess++
                                $reportEntry = [PSCustomObject]@{
                                    SiteName     = $siteUrl.Split("/")[-1]
                                    LibraryName  = $list.Title
                                    ItemID       = $item.Id
                                    UserUPN      = $UserUPN
                                    ItemType     = $itemType
                                    Name         = $itemName
                                    Location     = $itemLocation
                                    Size         = if ($itemSize) { "$([math]::Round($itemSize/1KB, 2)) KB" } else { "" }
                                    Read         = if ($readSources.Count -gt 0) { $readSources -join "`n" } else { "" }
                                    Edit         = if ($editSources.Count -gt 0) { $editSources -join "`n" } else { "" }
                                    FullControl  = if ($fullControl) { "Yes" } else { "" }
                                    UniquePerms  = if ($hasUniquePerms) { "Yes" } else { "No" }
                                    LastModified = if ($item.Modified) { $item.Modified } else { "" }
                                    SharingLinks = if ($sharingLinks.Count -gt 0) { $sharingLinks -join "`n" } else { "" }
                                }
                                $reportData += $reportEntry
                                
                                Write-Log "      Found access for $UserUPN to $itemType $itemName" -Level "INFO" -ForegroundColor "Green"
                                Write-Log "        Read Access: $($readSources.Count > 0)" -Level "DEBUG" -ForegroundColor "DarkGray"
                                Write-Log "        Edit Access: $($editSources.Count > 0)" -Level "DEBUG" -ForegroundColor "DarkGray"
                                Write-Log "        Full Control: $fullControl" -Level "DEBUG" -ForegroundColor "DarkGray"
                                Write-Log "        Unique Perms: $hasUniquePerms" -Level "DEBUG" -ForegroundColor "DarkGray"
                            }
                        }

                    } catch {
                        Write-Log "Error processing item $($item.Id): $_" -Level "ERROR" -ForegroundColor "Red"
                    }
                }
            } catch {
                Write-Log "Error retrieving items: $_" -Level "ERROR" -ForegroundColor "Red"
                $nextPageUrl = $null
            }
        } while ($nextPageUrl)

        $listDuration = (Get-Date) - $listStartTime
        Write-Log "Completed processing $($list.Title) in $($listDuration.TotalMinutes.ToString('0.00')) minutes" -Level "INFO" -ForegroundColor "Green"
    }

    # Generate CSV report
    if ($reportData.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $siteName = $siteUrl.Split("/")[-1]
        $fileName = "$siteName-UserAccessReport_$timestamp.csv"
        
        try {
            $reportData | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            Write-Log "`nSITE REPORT SUMMARY:" -Level "INFO" -ForegroundColor "Cyan"
            Write-Log "Total items processed: $siteItemsProcessed" -Level "INFO" -ForegroundColor "White"
            Write-Log "Items with matching access: $siteItemsWithAccess" -Level "INFO" -ForegroundColor "White"
            Write-Log "Report generated: $fileName" -Level "SUCCESS" -ForegroundColor "Green"
        }
        catch {
            Write-Log "Error generating report: $_" -Level "ERROR" -ForegroundColor "Red"
        }
    } else {
        Write-Log "`nNo matching permissions found for users in $siteUrl." -Level "INFO" -ForegroundColor "Yellow"
        Write-Log "Total items processed: $siteItemsProcessed" -Level "INFO" -ForegroundColor "White"
    }

    $siteDuration = (Get-Date) - $siteStartTime
    Write-Log "Completed processing site $siteUrl in $($siteDuration.TotalMinutes.ToString('0.00')) minutes" -Level "INFO" -ForegroundColor "Green"
    
    Disconnect-PnPOnline
    Write-Log "Disconnected from $siteUrl" -Level "INFO" -ForegroundColor "DarkGray"
}

# Final summary
$totalDuration = (Get-Date) - $global:StartTime
Write-Log "`nSCRIPT COMPLETED" -Level "INFO" -ForegroundColor "Cyan"
Write-Log "================================================================" -Level "HEADER" -ForegroundColor "Cyan"
Write-Log "TOTAL ITEMS PROCESSED: $global:TotalItemsProcessed" -Level "INFO" -ForegroundColor "White"
Write-Log "ITEMS WITH MATCHING ACCESS: $global:TotalItemsWithAccess" -Level "INFO" -ForegroundColor "White"
Write-Log "TOTAL EXECUTION TIME: $($totalDuration.TotalMinutes.ToString('0.00')) minutes" -Level "INFO" -ForegroundColor "White"
Write-Log "================================================================" -Level "HEADER" -ForegroundColor "Cyan"
