

$ExcludedLists = @("Access Requests", "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", 
    "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms",
    "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", 
    "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Images", "site collection images",
    "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", 
    "Reusable Content", "Reporting Metadata", "Reporting Templates", "Search Config List", "Site Assets", 
    "Preservation Hold Library", "Solution Gallery", "Style Library", "Suggested Content Browser Locations", 
    "Theme Gallery", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", 
    "Workflow History", "Workflow Tasks", "Pages")

# Function to properly escape CSV fields
function ConvertTo-CSVCell {
    param([string]$Value)
    
    if ($Value -eq $null -or $Value -eq "") {
        return ""
    }
    
    # Check if value contains comma, double quote, or newline
    if ($Value -match '[,"\n\r]') {
        # Escape double quotes by doubling them
        $Value = $Value -replace '"', '""'
        # Wrap in double quotes
        return "`"$Value`""
    }
    
    return $Value
}

# Create a master CSV file with timestamp
$masterTimestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$masterCsvFile = "SharePoint_Permissions_Report_$masterTimestamp.csv"

# Write CSV headers
$headers = @(
    "SiteName",
    "SiteUrl",
    "LibraryName",
    "ItemID",
    "ItemType",
    "Name",
    "Location",
    "SizeKB",
    "ReadUsers_WithLogin",
    "EditUsers_WithLogin",
    "FullControlUsers_WithLogin",
    "SharingLinks",
    "UniquePerms",
    "LastModified",
    "ProcessedAt"
)

$headerLine = ($headers | ForEach-Object { ConvertTo-CSVCell $_ }) -join ","
[System.IO.File]::WriteAllText($masterCsvFile, $headerLine + "`n")

Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "SHAREPOINT PERMISSIONS REPORT - Processing ALL items" -ForegroundColor Cyan
Write-Host "Output File: $masterCsvFile" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Cyan

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

    $processedItems = 0
    $itemsWritten = 0

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

        $nextPageUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items?`$top=1000&`$expand=File,Folder"
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
                    
                    try {
                        # Get item details and determine type
                        $fileSystemObjectType = $item.FileSystemObjectType
                        $itemName = $null
                        $itemLocation = $null
                        $itemSize = $null
                        
                        if ($fileSystemObjectType -eq 0 -and $item.File) {
                            $itemType = "File"
                            $itemName = $item.File.Name
                            $itemLocation = $item.File.ServerRelativeUrl
                            $itemSize = $item.File.Length
                            Write-Host "    Processing file: $itemName" -ForegroundColor DarkGray
                        } 
                        elseif ($fileSystemObjectType -eq 1 -and $item.Folder) {
                            $itemType = "Folder"
                            $itemName = $item.Folder.Name
                            $itemLocation = $item.Folder.ServerRelativeUrl
                            Write-Host "    Processing folder: $itemName" -ForegroundColor DarkGray
                        } 
                        else {
                            $itemType = "ListItem"
                            $itemName = if ($item.Title) { $item.Title } else { "Item_$($item.Id)" }
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

                        # Initialize permission collections
                        $readUsers = @()
                        $editUsers = @()
                        $fullControlUsers = @()
                        $sharingLinks = @()

                        # Get permissions info
                        try {
                            $permsInfo = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation" -Method Get -ErrorAction Stop
                            
                            # Process direct permissions and group memberships
                            if ($permsInfo.permissionsInformation.principals) {
                                foreach ($principalElement in $permsInfo.permissionsInformation.principals) {
                                    $principal = $principalElement.principal
                                    $role = $principalElement.role

                                    if ($principal.principalType -eq 1) {
                                        # User - get login name
                                        $userLogin = if ($principal.userPrincipalName) { $principal.userPrincipalName } else { $principal.email }
                                        $userName = if ($principal.name) { $principal.name } else { $userLogin }
                                        $userEntry = "$userName [$userLogin]"
                                        
                                        if ($hasUniquePerms) {
                                            $userEntry += " (UNIQUE)"
                                        } else {
                                            $userEntry += " (INHERITED)"
                                        }
                                        
                                        switch ($role) {
                                            1 { $readUsers += $userEntry }
                                            2 { $editUsers += $userEntry }
                                            3 { $fullControlUsers += $userEntry }
                                        }
                                    }
                                    elseif ($principal.principalType -in @(4,8)) {
                                        # Group - get members
                                        $groupName = $principal.name
                                        $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))/Users"
                                        try {
                                            $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                                            foreach ($member in $members.value) {
                                                if ($member.PrincipalType -eq 1) {
                                                    $memberLogin = if ($member.UserPrincipalName) { $member.UserPrincipalName } else { $member.Email }
                                                    $memberName = if ($member.Title) { $member.Title } else { $memberLogin }
                                                    $memberEntry = "$memberName [$memberLogin] (via $groupName)"
                                                    
                                                    if ($hasUniquePerms) {
                                                        $memberEntry += " (UNIQUE)"
                                                    } else {
                                                        $memberEntry += " (INHERITED)"
                                                    }
                                                    
                                                    switch ($role) {
                                                        1 { $readUsers += $memberEntry }
                                                        2 { $editUsers += $memberEntry }
                                                        3 { $fullControlUsers += $memberEntry }
                                                    }
                                                }
                                            }
                                        }
                                        catch {
                                            Write-Host "    Error getting members for group $groupName : $_" -ForegroundColor Yellow
                                            $groupEntry = "$groupName [Group Members Not Accessible]"
                                            
                                            if ($hasUniquePerms) {
                                                $groupEntry += " (UNIQUE)"
                                            } else {
                                                $groupEntry += " (INHERITED)"
                                            }
                                            
                                            switch ($role) {
                                                1 { $readUsers += $groupEntry }
                                                2 { $editUsers += $groupEntry }
                                                3 { $fullControlUsers += $groupEntry }
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
                                    $linkInfo = "$linkUrl ($linkType access)"
                                    
                                    if ($link.linkDetails.ExpirationDate) {
                                        $linkInfo += " - Expires: $($link.linkDetails.ExpirationDate)"
                                    }
                                    if ($link.linkDetails.PasswordProtected) {
                                        $linkInfo += " - Password Protected"
                                    }
                                    
                                    $sharingLinks += $linkInfo
                                    
                                    if ($link.linkMembers) {
                                        foreach ($member in $link.linkMembers) {
                                            $memberLogin = if ($member.userPrincipalName) { $member.userPrincipalName } else { $member.email }
                                            $memberName = if ($member.displayName) { $member.displayName } else { $memberLogin }
                                            $memberEntry = "$memberName [$memberLogin] (via sharing link)"
                                            
                                            if ($linkType -eq "Edit") {
                                                $editUsers += $memberEntry
                                            } else {
                                                $readUsers += $memberEntry
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch {
                            Write-Host "    Error getting permissions info for item $($item.Id): $_" -ForegroundColor Red
                        }

                        # If no users found, add default message
                        if ($readUsers.Count -eq 0 -and $editUsers.Count -eq 0 -and $fullControlUsers.Count -eq 0) {
                            $readUsers = @("No permissions found")
                        }
                        
                        # If no sharing links, add "None"
                        if ($sharingLinks.Count -eq 0) {
                            $sharingLinks = @("None")
                        }

                        # Prepare CSV row
                        $rowData = @(
                            $siteUrl.Split("/")[-1],
                            $siteUrl,
                            $list.Title,
                            $item.Id.ToString(),
                            $itemType,
                            $itemName,
                            $itemLocation,
                            $(if ($itemSize) { [math]::Round($itemSize/1KB, 2).ToString() } else { "" }),
                            ($readUsers | Sort-Object -Unique) -join "`n",
                            ($editUsers | Sort-Object -Unique) -join "`n",
                            ($fullControlUsers | Sort-Object -Unique) -join "`n",
                            ($sharingLinks | Sort-Object -Unique) -join "`n",
                            $(if ($hasUniquePerms) { "Yes - Unique" } else { "No - Inherited" }),
                            $(if ($item.Modified) { $item.Modified } else { "" }),
                            (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                        )
                        
                        # Escape each field and join with commas
                        $csvLine = ($rowData | ForEach-Object { ConvertTo-CSVCell $_ }) -join ","
                        
                        # Append to CSV file immediately (dynamic update)
                        Add-Content -Path $masterCsvFile -Value $csvLine -Encoding UTF8
                        $itemsWritten++
                        
                        # Show progress
                        $totalUsers = $readUsers.Count + $editUsers.Count + $fullControlUsers.Count
                        $permType = if ($hasUniquePerms) { "UNIQUE" } else { "INHERITED" }
                        
                        Write-Host "      ✅ [$itemsWritten] $permType - $itemType : $itemName (Users: $totalUsers, Links: $($sharingLinks.Count))" -ForegroundColor $(if ($hasUniquePerms) { "Green" } else { "DarkGray" })

                    } catch {
                        Write-Host "    ❌ Error processing item $($item.Id): $_" -ForegroundColor Red
                    }
                }
            } catch {
                Write-Host "  Error retrieving items: $_" -ForegroundColor Red
                $nextPageUrl = $null
            }
        } while ($nextPageUrl)
    }

    Write-Host "`n📁 Site Summary for $($siteUrl.Split("/")[-1]):" -ForegroundColor Yellow
    Write-Host "   Items processed: $processedItems" -ForegroundColor White
    Write-Host "   Items written to CSV: $itemsWritten" -ForegroundColor White
    
    $fileInfo = Get-Item $masterCsvFile -ErrorAction SilentlyContinue
    if ($fileInfo) {
        Write-Host "   Total CSV size: $([math]::Round($fileInfo.Length/1KB, 2)) KB" -ForegroundColor White
    }

    Disconnect-PnPOnline
    Write-Host "Disconnected from $siteUrl" -ForegroundColor DarkGray
}

Write-Host "`n====================================================================" -ForegroundColor Green
Write-Host "✅ REPORT GENERATION COMPLETED" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Green
Write-Host "Final Report: $masterCsvFile" -ForegroundColor Green

$finalFileInfo = Get-Item $masterCsvFile
$totalRows = (Get-Content $masterCsvFile | Measure-Object -Line).Lines - 1
Write-Host "Total rows (including header): $($totalRows + 1)" -ForegroundColor White
Write-Host "Data rows: $totalRows" -ForegroundColor White
Write-Host "File size: $([math]::Round($finalFileInfo.Length/1KB, 2)) KB" -ForegroundColor White
Write-Host "====================================================================" -ForegroundColor Green
Write-Host "✅ FEATURES INCLUDED:" -ForegroundColor Yellow
Write-Host "   • ALL items processed (both inherited and unique permissions)" -ForegroundColor White
Write-Host "   • User login names with display names" -ForegroundColor White
Write-Host "   • Group members expanded to individual users" -ForegroundColor White
Write-Host "   • Sharing links with expiration and password details" -ForegroundColor White
Write-Host "   • Proper CSV escaping (commas inside quotes)" -ForegroundColor White
Write-Host "   • Dynamic real-time updates to CSV file" -ForegroundColor White
Write-Host "   • Each user on new line within cell (enable Wrap Text in Excel)" -ForegroundColor White
Write-Host "====================================================================" -ForegroundColor Green

# Optional: Open the CSV file
# Invoke-Item $masterCsvFile
