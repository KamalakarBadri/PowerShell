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

# Loop through each SharePoint site
foreach ($siteUrl in $SiteUrls) {
    Write-Host "`n====================================================================" -ForegroundColor Cyan
    Write-Host "CONNECTING TO SITE: $siteUrl" -ForegroundColor Cyan
    Write-Host "====================================================================" -ForegroundColor Cyan
    
    try {
        Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId -ErrorAction Stop
        Write-Host "Successfully connected to $siteUrl" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to connect to $siteUrl : $_" -ForegroundColor Red
        continue
    }

    # Get all document libraries
    Write-Host "`nRetrieving document libraries..." -ForegroundColor Yellow
    try {
        $lists = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists" -Method Get -ErrorAction Stop
        Write-Host "Found $($lists.value.Count) lists in the site" -ForegroundColor Green
    }
    catch {
        Write-Host "Error retrieving lists: $_" -ForegroundColor Red
        continue
    }

    $reportData = @()
    $processedItems = 0
    $itemsWithAccess = 0

    foreach ($list in $lists.value) {
        if ($list.BaseTemplate -ne 101) { 
            Write-Host "Skipping non-document library: $($list.Title) (BaseTemplate: $($list.BaseTemplate))" -ForegroundColor Gray
            continue 
        }

        # Check if the list is in the exclusion list
        if ($ExcludedLists -contains $list.Title) {
            Write-Host "Skipping excluded list: $($list.Title)" -ForegroundColor Gray
            continue
        }

        Write-Host "`nProcessing Document Library: $($list.Title)" -ForegroundColor Yellow

        $nextPageUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items?`$top=1000"
        $pageCount = 0
        
        do {
            $pageCount++
            Write-Host "  Retrieving page $pageCount of items..." -ForegroundColor DarkGray
            
            try {
                $response = Invoke-PnPSPRestMethod -Url $nextPageUrl -Method Get -ErrorAction Stop
                $listItems = $response.value
                $nextPageUrl = $response."odata.nextLink"
                Write-Host "  Found $($listItems.Count) items on this page" -ForegroundColor DarkGray

                foreach ($item in $listItems) {
                    $processedItems++
                    if ($processedItems % 100 -eq 0) {
                        Write-Host "Processed $processedItems items total (found $itemsWithAccess with access so far)" -ForegroundColor DarkGray
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
                                Write-Host "    Processing file: $itemName" -ForegroundColor DarkGray
                            }
                            catch {
                                Write-Host "    Error getting file details for item $($item.Id): $_" -ForegroundColor Red
                                continue
                            }
                        } 
                        elseif ($fileSystemObjectType -eq 1) {
                            $itemType = "Folder"
                            try {
                                $folderResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/folder" -Method Get -ErrorAction Stop
                                $itemName = $folderResponse.Name
                                $itemLocation = $folderResponse.ServerRelativeUrl
                                Write-Host "    Processing folder: $itemName" -ForegroundColor DarkGray
                            }
                            catch {
                                Write-Host "    Error getting folder details for item $($item.Id): $_" -ForegroundColor Red
                                continue
                            }
                        } 
                        else {
                            $itemType = "ListItem"
                            $itemName = $item.Title
                            $itemLocation = $null
                            Write-Host "    Processing list item: $itemName" -ForegroundColor DarkGray
                        }

                        # Check if item has unique permissions
                        try {
                            $uniquePerms = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/HasUniqueRoleAssignments" -Method Get -ErrorAction Stop
                            $hasUniquePerms = $uniquePerms.value
                        }
                        catch {
                            Write-Host "    Error checking unique permissions for item $($item.Id): $_" -ForegroundColor Red
                            $hasUniquePerms = $false
                        }

                        # Get permissions info (both direct and inherited)
                        try {
                            $permsInfo = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation" -Method Get -ErrorAction Stop
                        }
                        catch {
                            Write-Host "    Error getting permissions info for item $($item.Id): $_" -ForegroundColor Red
                            continue
                        }

                        foreach ($UserUPN in $UserUPNs) {
                            $readSources = @()
                            $editSources = @()
                            $fullControl = $false
                            $sharingLinks = @()

                            # Process direct permissions and group memberships
                            if ($permsInfo.permissionsInformation.principals) {
                                Write-Host "    Found $($permsInfo.permissionsInformation.principals.Count) principals with permissions" -ForegroundColor DarkGray
                                
                                foreach ($principalElement in $permsInfo.permissionsInformation.principals) {
                                    $principal = $principalElement.principal
                                    $role = $principalElement.role
                                    Write-Host "      Processing principal: $($principal.name) (Type: $($principal.principalType)) with role: $role" -ForegroundColor DarkGray

                                    if ($principal.principalType -eq 1) {
                                        # User - check if it matches our target UPN
                                        $principalUpn = $principal.userPrincipalName ?? $principal.email
                                        $principalName = $principal.name ?? $principalUpn
                                        Write-Host "        User principal found: $principalName" -ForegroundColor DarkGray
                                        
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
                                        Write-Host "      Processing group: $groupName (Type: $($principal.principalType))" -ForegroundColor DarkGray
                                        
                                        $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))/Users"
                                        try {
                                            Write-Host "        Retrieving members for group $groupName..." -ForegroundColor DarkGray
                                            $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                                            Write-Host "        Found $($members.value.Count) members in group $groupName" -ForegroundColor DarkGray
                                            
                                            if ($members.value.Count -eq 0) {
                                                Write-Host "        No members found in group $groupName" -ForegroundColor Yellow
                                                continue
                                            }

                                            foreach ($member in $members.value) {
                                                Write-Host "          Processing member (Type: $($member.PrincipalType))..." -ForegroundColor DarkGray
                                                $memberName = $null
                                                $memberUpn = $null
                                                
                                                if ($member.PrincipalType -eq 1) {
                                                    # User
                                                    $memberUpn = $member.UserPrincipalName ?? $member.Email
                                                    $memberName = $member.Title ?? $memberUpn
                                                    Write-Host "            User member: $memberName" -ForegroundColor DarkGray
                                                }
                                                elseif ($member.PrincipalType -eq 4) {
                                                    # Security Group
                                                    $memberName = $member.Title ?? $member.LoginName
                                                    Write-Host "            Security Group member: $memberName" -ForegroundColor DarkGray
                                                }
                                                elseif ($member.PrincipalType -eq 8) {
                                                    # SharePoint Group
                                                    $memberName = $member.Title ?? $member.LoginName
                                                    Write-Host "            SharePoint Group member: $memberName" -ForegroundColor DarkGray
                                                }
                                                else {
                                                    # Other principal types (like 5 for distribution lists)
                                                    $memberName = $member.Title ?? $member.LoginName ?? $member.Email ?? "Unknown Principal"
                                                    Write-Host "            Other member type ($($member.PrincipalType)): $memberName" -ForegroundColor DarkGray
                                                }
                                                
                                                # If member is a user and matches our target UPN
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
                                            Write-Host "        Error getting members for group $groupName : $_" -ForegroundColor Yellow
                                            # If we can't get members, we can't verify access through this group
                                        }
                                    }
                                }
                            }

                            # Process sharing links
                            if ($permsInfo.permissionsInformation.links) {
                                Write-Host "    Found $($permsInfo.permissionsInformation.links.Count) sharing links" -ForegroundColor DarkGray
                                
                                foreach ($link in $permsInfo.permissionsInformation.links) {
                                    if ($link.linkDetails -and $link.linkDetails.Url) {
                                        $linkUrl = $link.linkDetails.Url
                                        $linkType = if ($link.linkDetails.IsEditLink -or $link.linkDetails.IsReviewLink) { "Edit" } else { "Read" }
                                        $sharingLinks += "$linkUrl ($linkType access)"
                                        Write-Host "      Found sharing link: $linkUrl ($linkType access)" -ForegroundColor DarkGray
                                        
                                        if ($link.linkMembers) {
                                            Write-Host "        Found $($link.linkMembers.Count) link members" -ForegroundColor DarkGray
                                            foreach ($member in $link.linkMembers) {
                                                $memberUpn = $member.userPrincipalName ?? $member.email
                                                $memberName = $member.displayName ?? $memberUpn
                                                Write-Host "          Link member: $memberName" -ForegroundColor DarkGray
                                                
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

                            # Only add to report if the user has Read, Edit, or Full Control
                            if ($readSources -or $editSources -or $fullControl) {
                                $itemsWithAccess++
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
                                
                                Write-Host "      Found access for $UserUPN to $itemType $itemName" -ForegroundColor Green
                                Write-Host "        Read Access: $($readSources.Count > 0)" -ForegroundColor DarkGray
                                Write-Host "        Edit Access: $($editSources.Count > 0)" -ForegroundColor DarkGray
                                Write-Host "        Full Control: $fullControl" -ForegroundColor DarkGray
                                Write-Host "        Unique Perms: $hasUniquePerms" -ForegroundColor DarkGray
                            }
                        }

                    } catch {
                        Write-Host "Error processing item $($item.Id): $_" -ForegroundColor Red
                    }
                }
            } catch {
                Write-Host "Error retrieving items: $_" -ForegroundColor Red
                $nextPageUrl = $null
            }
        } while ($nextPageUrl)
    }

    # Generate CSV report
    if ($reportData.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $siteName = $siteUrl.Split("/")[-1]
        $fileName = "$siteName-UserAccessReport_$timestamp.csv"
        
        try {
            $reportData | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            Write-Host "`nREPORT SUMMARY:" -ForegroundColor Cyan
            Write-Host "Total items processed: $processedItems" -ForegroundColor White
            Write-Host "Items with matching access: $itemsWithAccess" -ForegroundColor White
            Write-Host "Report generated: $fileName" -ForegroundColor Green
        }
        catch {
            Write-Host "Error generating report: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "`nNo matching permissions found for users in $siteUrl." -ForegroundColor Yellow
        Write-Host "Total items processed: $processedItems" -ForegroundColor White
    }

    Disconnect-PnPOnline
    Write-Host "Disconnected from $siteUrl" -ForegroundColor DarkGray
}

Write-Host "`nScript completed processing all sites." -ForegroundColor Cyan
