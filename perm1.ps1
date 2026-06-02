$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ThumbPrint = "B799789F78628CAE56B4D0F380FD551EB754E0DB"

# Array of site URLs to process
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

# Create timestamp and CSV file
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$csvFile = "PermissionsReport_$timestamp.csv"

Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "REPORT SAVED TO: $csvFile (updates in real-time)" -ForegroundColor Cyan
Write-Host "====================================================================" -ForegroundColor Cyan

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

    foreach ($list in $lists.value) {
        if ($list.BaseTemplate -ne 101) { continue }
        if ($ExcludedLists -contains $list.Title) { 
            Write-Host "Skipping excluded list: $($list.Title)" -ForegroundColor Gray
            continue 
        }

        Write-Host "`nProcessing Document Library: $($list.Title)" -ForegroundColor Yellow

        $nextPageUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items?`$top=1000"
        
        do {
            try {
                $response = Invoke-PnPSPRestMethod -Url $nextPageUrl -Method Get -ErrorAction Stop
                $listItems = $response.value
                $nextPageUrl = $response."odata.nextLink"

                foreach ($item in $listItems) {
                    $processedItems++
                    
                    if ($processedItems % 100 -eq 0) {
                        Write-Host "Processed $processedItems items (found $itemsWithPermissions with permissions)" -ForegroundColor DarkGray
                    }

                    try {
                        # Get item details
                        $fileSystemObjectType = $item.FileSystemObjectType
                        $itemName = $null
                        $itemLocation = $null
                        $itemSize = $null
                        
                        if ($fileSystemObjectType -eq 0) {
                            $itemType = "File"
                            $fileResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/file" -Method Get -ErrorAction Stop
                            $itemName = $fileResponse.Name
                            $itemLocation = $fileResponse.ServerRelativeUrl
                            $itemSize = $fileResponse.Length
                        } 
                        elseif ($fileSystemObjectType -eq 1) {
                            $itemType = "Folder"
                            $folderResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/folder" -Method Get -ErrorAction Stop
                            $itemName = $folderResponse.Name
                            $itemLocation = $folderResponse.ServerRelativeUrl
                        } 
                        else {
                            $itemType = "ListItem"
                            $itemName = $item.Title
                            $itemLocation = $null
                        }

                        # Check unique permissions
                        $uniquePerms = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/HasUniqueRoleAssignments" -Method Get -ErrorAction Stop
                        $hasUniquePerms = $uniquePerms.value

                        # Get permissions info
                        $permsInfo = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation" -Method Get -ErrorAction Stop

                        $readUsers = @()
                        $editUsers = @()
                        $fullControlUsers = @()
                        $sharingLinks = @()

                        # Process permissions
                        if ($permsInfo.permissionsInformation.principals) {
                            foreach ($principalElement in $permsInfo.permissionsInformation.principals) {
                                $principal = $principalElement.principal
                                $role = $principalElement.role

                                if ($principal.principalType -eq 1) {
                                    # User - get both display name and login name
                                    $principalLogin = $principal.userPrincipalName ?? $principal.email ?? $principal.loginName
                                    $principalDisplayName = $principal.name ?? $principal.title ?? $principalLogin
                                    
                                    # Format as "Display Name (Login Name)"
                                    $principalFormatted = "$principalDisplayName ($principalLogin)"
                                    
                                    switch ($role) {
                                        1 { $readUsers += $principalFormatted }
                                        2 { $editUsers += $principalFormatted }
                                        3 { $fullControlUsers += $principalFormatted }
                                    }
                                }
                                elseif ($principal.principalType -in @(4,8)) {
                                    # Group - get ALL members with their login names
                                    $groupName = $principal.name
                                    $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))/Users"
                                    try {
                                        $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                                        foreach ($member in $members.value) {
                                            if ($member.PrincipalType -eq 1) {
                                                $memberLogin = $member.UserPrincipalName ?? $member.Email ?? $member.LoginName
                                                $memberDisplayName = $member.Title ?? $memberLogin
                                                
                                                # Format as "Display Name (Login Name) via GroupName"
                                                $memberFormatted = "$memberDisplayName ($memberLogin) via $groupName"
                                                
                                                switch ($role) {
                                                    1 { $readUsers += $memberFormatted }
                                                    2 { $editUsers += $memberFormatted }
                                                    3 { $fullControlUsers += $memberFormatted }
                                                }
                                            }
                                        }
                                    }
                                    catch {
                                        # If we can't get members, just add the group name
                                        $groupFormatted = "$groupName [members not accessible]"
                                        switch ($role) {
                                            1 { $readUsers += $groupFormatted }
                                            2 { $editUsers += $groupFormatted }
                                            3 { $fullControlUsers += $groupFormatted }
                                        }
                                    }
                                }
                            }
                        }

                        # Process sharing links - FIXED: Only add if actual sharing links exist
                        if ($permsInfo.permissionsInformation.links -and $permsInfo.permissionsInformation.links.Count -gt 0) {
                            foreach ($link in $permsInfo.permissionsInformation.links) {
                                # Skip if it's a default "Read Access" link without proper sharing link properties
                                $isValidSharingLink = $false
                                
                                # Check if this is a real sharing link (has sharing link properties)
                                if ($link.linkDetails -and $link.linkDetails.Url) {
                                    # Check if it's an actual sharing link (contains 'guestinvite' or 'sharing' or has unique ID)
                                    $linkUrl = $link.linkDetails.Url
                                    if ($linkUrl -match "guestinvite|sharing|/SharedWith/|sharedlink") {
                                        $isValidSharingLink = $true
                                    }
                                    # Also check if it has sharing token or link ID
                                    elseif ($link.linkDetails.SharingToken -or $link.linkDetails.LinkId) {
                                        $isValidSharingLink = $true
                                    }
                                    # If we have link members, it's definitely a valid sharing link
                                    elseif ($link.linkMembers -and $link.linkMembers.Count -gt 0) {
                                        $isValidSharingLink = $true
                                    }
                                }
                                
                                if ($isValidSharingLink) {
                                    $linkUrl = $link.linkDetails.Url
                                    $linkType = if ($link.linkDetails.IsEditLink -or $link.linkDetails.IsReviewLink) { "Edit" } else { "Read" }
                                    $sharingLinks += "$linkUrl ($linkType access)"
                                    
                                    if ($link.linkMembers) {
                                        foreach ($member in $link.linkMembers) {
                                            $memberLogin = $member.userPrincipalName ?? $member.email ?? $member.loginName
                                            $memberDisplayName = $member.displayName ?? $member.name ?? $memberLogin
                                            
                                            # Format as "Display Name (Login Name) via sharing link"
                                            $memberFormatted = "$memberDisplayName ($memberLogin) via sharing link: $linkUrl"
                                            
                                            if ($linkType -eq "Edit") {
                                                $editUsers += $memberFormatted
                                            } else {
                                                $readUsers += $memberFormatted
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        # Create report entry
                        $itemsWithPermissions++
                        $reportEntry = [PSCustomObject]@{
                            SiteName         = $siteUrl.Split("/")[-1]
                            LibraryName      = $list.Title
                            ItemID           = $item.Id
                            ItemType         = $itemType
                            Name             = $itemName
                            Location         = $itemLocation
                            Size             = if ($itemSize) { "$([math]::Round($itemSize/1KB, 2)) KB" } else { "" }
                            ReadUsers        = ($readUsers | Sort-Object -Unique) -join "`n"
                            EditUsers        = ($editUsers | Sort-Object -Unique) -join "`n"
                            FullControlUsers = ($fullControlUsers | Sort-Object -Unique) -join "`n"
                            SharingLinks     = ($sharingLinks | Sort-Object -Unique) -join "`n"
                            UniquePerms      = if ($hasUniquePerms) { "Yes" } else { "No" }
                            LastModified     = if ($item.Modified) { $item.Modified } else { "" }
                        }
                        
                        # Write to CSV dynamically (append mode)
                        if ($processedItems -eq 1 -and $itemsWithPermissions -eq 1) {
                            # First item - create file with header
                            $reportEntry | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
                        } else {
                            # Subsequent items - append to existing file
                            $reportEntry | Export-Csv -Path $csvFile -Append -NoTypeInformation -Encoding UTF8
                        }
                        
                        Write-Host "  Found permissions for $itemType: $itemName" -ForegroundColor Green
                    } 
                    catch {
                        Write-Host "  Error processing item $($item.Id): $_" -ForegroundColor Red
                    }
                }
            } 
            catch {
                Write-Host "Error retrieving items: $_" -ForegroundColor Red
                $nextPageUrl = $null
            }
        } while ($nextPageUrl)
    }

    Write-Host "`nSITE SUMMARY:" -ForegroundColor Cyan
    Write-Host "Total items processed: $processedItems" -ForegroundColor White
    Write-Host "Items with permissions recorded: $itemsWithPermissions" -ForegroundColor White
    Write-Host "Report updated in: $csvFile" -ForegroundColor Green

    Disconnect-PnPOnline
}

Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "SCRIPT COMPLETED - Report saved to: $csvFile" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Cyan
