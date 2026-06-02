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
                                    $principalUpn = $principal.userPrincipalName ?? $principal.email
                                    $principalName = $principal.name ?? $principalUpn
                                    
                                    switch ($role) {
                                        1 { $readUsers += $principalName }
                                        2 { $editUsers += $principalName }
                                        3 { $fullControlUsers += $principalName }
                                    }
                                }
                                elseif ($principal.principalType -in @(4,8)) {
                                    $groupName = $principal.name
                                    $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))/Users"
                                    try {
                                        $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                                        foreach ($member in $members.value) {
                                            if ($member.PrincipalType -eq 1) {
                                                $memberUpn = $member.UserPrincipalName ?? $member.Email
                                                $memberName = $member.Title ?? $memberUpn
                                                
                                                switch ($role) {
                                                    1 { $readUsers += "$memberName (via $groupName)" }
                                                    2 { $editUsers += "$memberName (via $groupName)" }
                                                    3 { $fullControlUsers += "$memberName (via $groupName)" }
                                                }
                                            }
                                        }
                                    }
                                    catch {
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
