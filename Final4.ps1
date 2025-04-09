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

# Function to recursively get all members of a group including nested groups
function Get-AllGroupMembers {
    param (
        [string]$siteUrl,
        [string]$groupId,
        [int]$depth = 0,
        [int]$maxDepth = 10
    )
    
    $members = @()
    
    if ($depth -ge $maxDepth) {
        Write-Host "      Maximum recursion depth ($maxDepth) reached for group $groupId" -ForegroundColor Yellow
        return $members
    }
    
    $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($groupId)/Users"
    try {
        $groupMembers = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
        
        foreach ($member in $groupMembers.value) {
            if ($member.PrincipalType -eq 1) {
                # User
                $memberUpn = $member.UserPrincipalName ?? $member.Email
                $memberName = $member.Title ?? $memberUpn
                $members += [PSCustomObject]@{
                    Name = $memberName
                    Type = "User"
                    Path = ""
                }
            }
            elseif ($member.PrincipalType -in @(4,8)) {
                # Nested Group - recursively get members
                $nestedGroupName = $member.Title
                Write-Host "      Found nested group: $nestedGroupName (depth $depth)" -ForegroundColor DarkGray
                $nestedMembers = Get-AllGroupMembers -siteUrl $siteUrl -groupId $member.Id -depth ($depth + 1) -maxDepth $maxDepth
                
                foreach ($nestedMember in $nestedMembers) {
                    $members += [PSCustomObject]@{
                        Name = $nestedMember.Name
                        Type = $nestedMember.Type
                        Path = "$nestedGroupName > $($nestedMember.Path)"
                    }
                }
            }
        }
    }
    catch {
        Write-Host "      Error getting members for group $groupId : $_" -ForegroundColor Yellow
    }
    
    return $members
}

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
    $itemsWithPermissions = 0

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
                                    # Group - recursively get ALL members including nested groups
                                    $groupName = $principal.name
                                    Write-Host "      Processing group: $groupName" -ForegroundColor DarkGray
                                    
                                    try {
                                        $allMembers = Get-AllGroupMembers -siteUrl $siteUrl -groupId $principal.id
                                        
                                        foreach ($member in $allMembers) {
                                            $displayName = $member.Name
                                            if ($member.Path) {
                                                $displayName = "$displayName (via $($member.Path))"
                                            } else {
                                                $displayName = "$displayName (via $groupName)"
                                            }
                                            
                                            switch ($role) {
                                                1 { $readUsers += $displayName }
                                                2 { $editUsers += $displayName }
                                                3 { $fullControlUsers += $displayName }
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

                        # Process sharing links - Updated to handle empty links
                        if ($permsInfo.permissionsInformation.links) {
                            foreach ($link in $permsInfo.permissionsInformation.links) {
                                if ($link.linkDetails -and $link.linkDetails.Url) {
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
                        }

                        # Add to report if there are any permissions (even if inherited)
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
                            
                            Write-Host "      Found permissions for $itemType $itemName" -ForegroundColor Green
                            Write-Host "        Read Users: $($readUsers.Count)" -ForegroundColor DarkGray
                            Write-Host "        Edit Users: $($editUsers.Count)" -ForegroundColor DarkGray
                            Write-Host "        Full Control Users: $($fullControlUsers.Count)" -ForegroundColor DarkGray
                            Write-Host "        Sharing Links: $($sharingLinks.Count)" -ForegroundColor DarkGray
                            Write-Host "        Unique Perms: $hasUniquePerms" -ForegroundColor DarkGray
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
        $fileName = "$siteName-PermissionsReport_$timestamp.csv"
        
        try {
            $reportData | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            Write-Host "`nREPORT SUMMARY:" -ForegroundColor Cyan
            Write-Host "Total items processed: $processedItems" -ForegroundColor White
            Write-Host "Items with permissions recorded: $itemsWithPermissions" -ForegroundColor White
            Write-Host "Report generated: $fileName" -ForegroundColor Green
        }
        catch {
            Write-Host "Error generating report: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "`nNo permissions found for items in $siteUrl." -ForegroundColor Yellow
        Write-Host "Total items processed: $processedItems" -ForegroundColor White
    }

    Disconnect-PnPOnline
    Write-Host "Disconnected from $siteUrl" -ForegroundColor DarkGray
}

Write-Host "`nScript completed processing all sites." -ForegroundColor Cyan
