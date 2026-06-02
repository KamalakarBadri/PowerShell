

$ExcludedLists = @("Access Requests", "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", 
    "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms",
    "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", 
    "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Images", "site collection images",
    "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", 
    "Reusable Content", "Reporting Metadata", "Reporting Templates", "Search Config List", "Site Assets", 
    "Preservation Hold Library", "Solution Gallery", "Style Library", "Suggested Content Browser Locations", 
    "Theme Gallery", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", 
    "Workflow History", "Workflow Tasks", "Pages")

# Create a master CSV file that will be updated throughout the script
$masterTimestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$masterFileName = "Master-PermissionsReport_$masterTimestamp.csv"

# Create CSV header if file doesn't exist
$header = "SiteName,LibraryName,ItemID,ItemType,Name,Location,Size,ReadUsers,EditUsers,FullControlUsers,SharingLinks,UniquePerms,LastModified`n"
[System.IO.File]::WriteAllText($masterFileName, $header)

Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "MASTER REPORT WILL BE SAVED TO: $masterFileName" -ForegroundColor Cyan
Write-Host "====================================================================" -ForegroundColor Cyan

# Loop through each SharePoint site
foreach ($siteUrl in $SiteUrls) {
    Write-Host "`n====================================================================" -ForegroundColor Cyan
    Write-Host "CONNECTING TO SITE: $siteUrl" -ForegroundColor Cyan
    Write-Host "====================================================================" -ForegroundColor Cyan
    
    # Create site-specific CSV file (optional - for separate reports per site)
    $siteSpecificFile = "$($siteUrl.Split('/')[-1])-PermissionsReport_Dynamic.csv"
    $siteHeader = "SiteName,LibraryName,ItemID,ItemType,Name,Location,Size,ReadUsers,EditUsers,FullControlUsers,SharingLinks,UniquePerms,LastModified`n"
    [System.IO.File]::WriteAllText($siteSpecificFile, $siteHeader)
    Write-Host "Site-specific report will be saved to: $siteSpecificFile" -ForegroundColor Green
    
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

    $processedItems = 0
    $itemsWithPermissions = 0
    $batchSize = 10  # Write to CSV after every N items (adjust as needed)
    $batchCounter = 0
    $batchData = @()

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
                        Write-Host "Processed $processedItems items total (found $itemsWithPermissions with permissions so far)" -ForegroundColor DarkGray
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

                        # Initialize permission collections
                        $readUsers = @()
                        $editUsers = @()
                        $fullControlUsers = @()
                        $sharingLinks = @()

                        # Process direct permissions and group memberships
                        if ($permsInfo.permissionsInformation.principals) {
                            foreach ($principalElement in $permsInfo.permissionsInformation.principals) {
                                $principal = $principalElement.principal
                                $role = $principalElement.role

                                if ($principal.principalType -eq 1) {
                                    # User - add to appropriate permission collection
                                    $principalUpn = $principal.userPrincipalName ?? $principal.email
                                    $principalName = $principal.name ?? $principalUpn
                                    
                                    switch ($role) {
                                        1 { $readUsers += $principalName }
                                        2 { $editUsers += $principalName }
                                        3 { $fullControlUsers += $principalName }
                                    }
                                }
                                elseif ($principal.principalType -in @(4,8)) {
                                    # Group - get ALL members and add them with group name
                                    $groupName = $principal.name
                                    $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))/Users"
                                    try {
                                        $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                                        foreach ($member in $members.value) {
                                            if ($member.PrincipalType -eq 1) {
                                                $memberUpn = $member.UserPrincipalName ?? $member.Email
                                                $memberName = $member.Title ?? $memberUpn
                                                
                                                # Add ALL group members to appropriate permission collection with group name
                                                switch ($role) {
                                                    1 { $readUsers += "$memberName (via $groupName)" }
                                                    2 { $editUsers += "$memberName (via $groupName)" }
                                                    3 { $fullControlUsers += "$memberName (via $groupName)" }
                                                }
                                            }
                                        }
                                    }
                                    catch {
                                        Write-Host "    Error getting members for group $groupName : $_" -ForegroundColor Yellow
                                        # If we can't get members, just add the group name
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
                            foreach ($link in $permsInfo.permissionsInformation.links) {
                                $linkUrl = $link.linkDetails.Url
                                $linkType = if ($link.linkDetails.IsEditLink -or $link.linkDetails.IsReviewLink) { "Edit" } else { "Read" }
                                $sharingLinks += "$linkUrl ($linkType access)"
                                
                                if ($link.linkMembers) {
                                    foreach ($member in $link.linkMembers) {
                                        $memberUpn = $member.userPrincipalName ?? $member.email
                                        $memberName = $member.displayName ?? $memberUpn
                                        
                                        if ($linkType -eq "Edit") {
                                            $editUsers += "$memberName (via sharing link : $linkUrl)"
                                        } else {
                                            $readUsers += "$memberName (via sharing link : $linkUrl)"
                                        }
                                    }
                                }
                            }
                        }

                        # Add to report if there are any permissions (even if inherited)
                        $itemsWithPermissions++
                        
                        # Escape special characters for CSV (handle commas, newlines, quotes)
                        $escapedReadUsers = ($readUsers | Sort-Object -Unique) -join "`n" | ForEach-Object { $_ -replace '"', '""' }
                        $escapedEditUsers = ($editUsers | Sort-Object -Unique) -join "`n" | ForEach-Object { $_ -replace '"', '""' }
                        $escapedFullControlUsers = ($fullControlUsers | Sort-Object -Unique) -join "`n" | ForEach-Object { $_ -replace '"', '""' }
                        $escapedSharingLinks = ($sharingLinks | Sort-Object -Unique) -join "`n" | ForEach-Object { $_ -replace '"', '""' }
                        
                        # Create CSV row (properly quoted)
                        $csvRow = @(
                            "`"$($siteUrl.Split('/')[-1])`"",
                            "`"$($list.Title -replace '"', '""')`"",
                            $item.Id,
                            "`"$itemType`"",
                            "`"$($itemName -replace '"', '""')`"",
                            "`"$($itemLocation -replace '"', '""')`"",
                            "`"$(if ($itemSize) { "$([math]::Round($itemSize/1KB, 2)) KB" } else { "" })`"",
                            "`"$escapedReadUsers`"",
                            "`"$escapedEditUsers`"",
                            "`"$escapedFullControlUsers`"",
                            "`"$escapedSharingLinks`"",
                            "`"$(if ($hasUniquePerms) { "Yes" } else { "No" })`"",
                            "`"$(if ($item.Modified) { $item.Modified } else { "" })`""
                        ) -join ","
                        
                        # Add to batch
                        $batchData += $csvRow
                        $batchCounter++
                        
                        # Write to CSV files in batches
                        if ($batchCounter -ge $batchSize) {
                            # Append to master CSV
                            [System.IO.File]::AppendAllText($masterFileName, ($batchData -join "`n") + "`n")
                            # Append to site-specific CSV
                            [System.IO.File]::AppendAllText($siteSpecificFile, ($batchData -join "`n") + "`n")
                            
                            Write-Host "      [DYNAMIC UPDATE] Wrote $batchCounter items to CSV files" -ForegroundColor Magenta
                            $batchData = @()
                            $batchCounter = 0
                        }
                        
                        Write-Host "      Found permissions for $itemType $itemName" -ForegroundColor Green
                        Write-Host "        Read Users: $($readUsers.Count)" -ForegroundColor DarkGray
                        Write-Host "        Edit Users: $($editUsers.Count)" -ForegroundColor DarkGray
                        Write-Host "        Full Control Users: $($fullControlUsers.Count)" -ForegroundColor DarkGray
                        Write-Host "        Sharing Links: $($sharingLinks.Count)" -ForegroundColor DarkGray
                        Write-Host "        Unique Perms: $hasUniquePerms" -ForegroundColor DarkGray

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

    # Write any remaining items in batch
    if ($batchData.Count -gt 0) {
        [System.IO.File]::AppendAllText($masterFileName, ($batchData -join "`n") + "`n")
        [System.IO.File]::AppendAllText($siteSpecificFile, ($batchData -join "`n") + "`n")
        Write-Host "`n[DYNAMIC UPDATE] Wrote final $($batchData.Count) items to CSV files" -ForegroundColor Magenta
    }

    # Summary for the site
    Write-Host "`nSITE SUMMARY FOR: $siteUrl" -ForegroundColor Cyan
    Write-Host "Total items processed: $processedItems" -ForegroundColor White
    Write-Host "Items with permissions recorded: $itemsWithPermissions" -ForegroundColor White
    Write-Host "Site-specific report: $siteSpecificFile" -ForegroundColor Green
    Write-Host "Updated in master report: $masterFileName" -ForegroundColor Green

    Disconnect-PnPOnline
    Write-Host "Disconnected from $siteUrl" -ForegroundColor DarkGray
}

Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "SCRIPT COMPLETED" -ForegroundColor Cyan
Write-Host "Master report saved to: $masterFileName" -ForegroundColor Green
Write-Host "You can open this file in Excel to view the live updated results" -ForegroundColor Yellow
Write-Host "====================================================================" -ForegroundColor Cyan
