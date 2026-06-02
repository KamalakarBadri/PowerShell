
$ExcludedLists = @("Access Requests", "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", 
    "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms",
    "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", 
    "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Images", "site collection images",
    "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", 
    "Reusable Content", "Reporting Metadata", "Reporting Templates", "Search Config List", "Site Assets", 
    "Preservation Hold Library", "Solution Gallery", "Style Library", "Suggested Content Browser Locations", 
    "Theme Gallery", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", 
    "Workflow History", "Workflow Tasks", "Pages")

# Create a master CSV file with timestamp
$masterTimestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$masterCsvFile = "Master_PermissionsReport_$masterTimestamp.csv"

# Function to properly escape CSV fields - CRITICAL for handling commas
function ConvertTo-CSVCell {
    param([string]$Value)
    
    if ($Value -eq $null -or $Value -eq "") {
        return ""
    }
    
    # Replace newlines with literal \n for display, or keep as is
    # Check if value contains comma, double quote, or newline
    if ($Value -match '[,"\n\r]') {
        # Escape double quotes by doubling them
        $Value = $Value -replace '"', '""'
        # Wrap in double quotes
        return "`"$Value`""
    }
    
    return $Value
}

# Write CSV header immediately
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

# Write header row
$headerLine = ($headers | ForEach-Object { ConvertTo-CSVCell $_ }) -join ","
[System.IO.File]::WriteAllText($masterCsvFile, $headerLine + "`n")

Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "DYNAMIC CSV REPORT - Updates in real-time" -ForegroundColor Cyan
Write-Host "Headers: $($headers -join ', ')" -ForegroundColor Green
Write-Host "Output File: $masterCsvFile" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Cyan

# Function to get user login names and display names
function Get-UserDetails {
    param($UserId, $SiteUrl)
    
    try {
        $userUrl = "$SiteUrl/_api/web/getuserbyid($UserId)"
        $user = Invoke-PnPSPRestMethod -Url $userUrl -Method Get -ErrorAction Stop
        return @{
            LoginName = $user.LoginName
            DisplayName = $user.Title
            Email = $user.Email
        }
    }
    catch {
        return @{
            LoginName = "Unknown"
            DisplayName = "Unknown"
            Email = ""
        }
    }
}

# Function to get sharing links for an item
function Get-SharingLinks {
    param($ListId, $ItemId, $SiteUrl)
    
    $sharingLinks = @()
    
    try {
        $sharingUrl = "$SiteUrl/_api/web/lists(guid'$ListId')/items($ItemId)/GetSharingInformation"
        $sharingInfo = Invoke-PnPSPRestMethod -Url $sharingUrl -Method Get -ErrorAction Stop
        
        if ($sharingInfo.HasSharingLinks -and $sharingInfo.SharingLinks) {
            foreach ($link in $sharingInfo.SharingLinks) {
                $linkInfo = "$($link.LinkUrl) [$($link.ShareType) access"
                if ($link.ExpirationDate) {
                    $linkInfo += " | Expires: $($link.ExpirationDate)"
                }
                if ($link.PasswordProtected) {
                    $linkInfo += " | Password Protected"
                }
                $linkInfo += "]"
                $sharingLinks += $linkInfo
            }
        }
    }
    catch {
        # No sharing links found
    }
    
    if ($sharingLinks.Count -eq 0) {
        return "None"
    }
    
    return ($sharingLinks -join "`n")
}

# Function to get permissions for an item (both direct and inherited)
function Get-ItemPermissions {
    param($ListId, $ItemId, $SiteUrl)
    
    $readUsers = @()
    $editUsers = @()
    $fullControlUsers = @()
    $hasUniquePerms = $false
    
    try {
        # Check if item has unique permissions
        $uniquePermsCheck = Invoke-PnPSPRestMethod -Url "$SiteUrl/_api/web/lists(guid'$ListId')/items($ItemId)/HasUniqueRoleAssignments" -Method Get -ErrorAction Stop
        $hasUniquePerms = $uniquePermsCheck.value
        
        if ($hasUniquePerms) {
            # Get unique permissions for this item
            $roleAssignmentsUrl = "$SiteUrl/_api/web/lists(guid'$ListId')/items($ItemId)/roleassignments?`$expand=Member,RoleDefinitionBindings"
            $roleAssignments = Invoke-PnPSPRestMethod -Url $roleAssignmentsUrl -Method Get -ErrorAction Stop
            
            foreach ($ra in $roleAssignments.value) {
                $member = $ra.Member
                $roleBindings = $ra.RoleDefinitionBindings
                
                $permissionLevel = ""
                foreach ($role in $roleBindings) {
                    if ($role.Name -eq "Read" -or $role.Name -eq "Restricted Read") {
                        $permissionLevel = "Read"
                    }
                    elseif ($role.Name -eq "Contribute" -or $role.Name -eq "Edit") {
                        $permissionLevel = "Edit"
                    }
                    elseif ($role.Name -eq "Full Control" -or $role.Name -eq "Owner") {
                        $permissionLevel = "FullControl"
                    }
                }
                
                if ($member.PrincipalType -eq 1) {
                    $userDetails = Get-UserDetails -UserId $member.Id -SiteUrl $SiteUrl
                    $userEntry = "$($userDetails.DisplayName) [$($userDetails.LoginName)]"
                    
                    switch ($permissionLevel) {
                        "Read" { $readUsers += $userEntry }
                        "Edit" { $editUsers += $userEntry }
                        "FullControl" { $fullControlUsers += $userEntry }
                    }
                }
                elseif ($member.PrincipalType -eq 4 -or $member.PrincipalType -eq 8) {
                    $groupName = $member.Title
                    $groupMembersUrl = "$SiteUrl/_api/web/SiteGroups/GetById($($member.Id))/Users"
                    try {
                        $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                        foreach ($groupMember in $members.value) {
                            $userDetails = Get-UserDetails -UserId $groupMember.Id -SiteUrl $SiteUrl
                            $userEntry = "$($userDetails.DisplayName) [$($userDetails.LoginName)] (via $groupName - UNIQUE)"
                            
                            switch ($permissionLevel) {
                                "Read" { $readUsers += $userEntry }
                                "Edit" { $editUsers += $userEntry }
                                "FullControl" { $fullControlUsers += $userEntry }
                            }
                        }
                    }
                    catch {
                        $groupEntry = "$groupName [Group Members Not Accessible] (UNIQUE)"
                        switch ($permissionLevel) {
                            "Read" { $readUsers += $groupEntry }
                            "Edit" { $editUsers += $groupEntry }
                            "FullControl" { $fullControlUsers += $groupEntry }
                        }
                    }
                }
            }
        }
        else {
            # Get inherited permissions from the library
            $listRoleAssignmentsUrl = "$SiteUrl/_api/web/lists(guid'$ListId')/roleassignments?`$expand=Member,RoleDefinitionBindings"
            $listRoleAssignments = Invoke-PnPSPRestMethod -Url $listRoleAssignmentsUrl -Method Get -ErrorAction Stop
            
            foreach ($ra in $listRoleAssignments.value) {
                $member = $ra.Member
                $roleBindings = $ra.RoleDefinitionBindings
                
                $permissionLevel = ""
                foreach ($role in $roleBindings) {
                    if ($role.Name -eq "Read" -or $role.Name -eq "Restricted Read") {
                        $permissionLevel = "Read"
                    }
                    elseif ($role.Name -eq "Contribute" -or $role.Name -eq "Edit") {
                        $permissionLevel = "Edit"
                    }
                    elseif ($role.Name -eq "Full Control" -or $role.Name -eq "Owner") {
                        $permissionLevel = "FullControl"
                    }
                }
                
                if ($member.PrincipalType -eq 1) {
                    $userDetails = Get-UserDetails -UserId $member.Id -SiteUrl $SiteUrl
                    $userEntry = "$($userDetails.DisplayName) [$($userDetails.LoginName)] (INHERITED)"
                    
                    switch ($permissionLevel) {
                        "Read" { $readUsers += $userEntry }
                        "Edit" { $editUsers += $userEntry }
                        "FullControl" { $fullControlUsers += $userEntry }
                    }
                }
                elseif ($member.PrincipalType -eq 4 -or $member.PrincipalType -eq 8) {
                    $groupName = $member.Title
                    $groupMembersUrl = "$SiteUrl/_api/web/SiteGroups/GetById($($member.Id))/Users"
                    try {
                        $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                        foreach ($groupMember in $members.value) {
                            $userDetails = Get-UserDetails -UserId $groupMember.Id -SiteUrl $SiteUrl
                            $userEntry = "$($userDetails.DisplayName) [$($userDetails.LoginName)] (via $groupName - INHERITED)"
                            
                            switch ($permissionLevel) {
                                "Read" { $readUsers += $userEntry }
                                "Edit" { $editUsers += $userEntry }
                                "FullControl" { $fullControlUsers += $userEntry }
                            }
                        }
                    }
                    catch {
                        $groupEntry = "$groupName [Group Members Not Accessible] (INHERITED)"
                        switch ($permissionLevel) {
                            "Read" { $readUsers += $groupEntry }
                            "Edit" { $editUsers += $groupEntry }
                            "FullControl" { $fullControlUsers += $groupEntry }
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Host "      Error getting permissions: $_" -ForegroundColor Red
    }
    
    # If no users found, add default message
    if ($readUsers.Count -eq 0 -and $editUsers.Count -eq 0 -and $fullControlUsers.Count -eq 0) {
        $readUsers = @("No permissions found")
    }
    
    return @{
        ReadUsers = $readUsers | Sort-Object -Unique
        EditUsers = $editUsers | Sort-Object -Unique
        FullControlUsers = $fullControlUsers | Sort-Object -Unique
        HasUniquePerms = $hasUniquePerms
    }
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

    $processedItems = 0
    $itemsWritten = 0

    foreach ($list in $lists.value) {
        if ($list.BaseTemplate -ne 101) { 
            Write-Host "Skipping non-document library: $($list.Title)" -ForegroundColor Gray
            continue 
        }

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
                        # Get item details
                        $itemName = ""
                        $itemLocation = ""
                        $itemSize = ""
                        $itemType = ""
                        
                        if ($item.FileSystemObjectType -eq 0 -and $item.File) {
                            $itemType = "File"
                            $itemName = $item.File.Name
                            $itemLocation = $item.File.ServerRelativeUrl
                            $itemSize = [math]::Round($item.File.Length/1KB, 2).ToString()
                        } 
                        elseif ($item.FileSystemObjectType -eq 1 -and $item.Folder) {
                            $itemType = "Folder"
                            $itemName = $item.Folder.Name
                            $itemLocation = $item.Folder.ServerRelativeUrl
                            $itemSize = ""
                        } 
                        else {
                            $itemType = "ListItem"
                            $itemName = if ($item.Title) { $item.Title } else { "Item_$($item.Id)" }
                            $itemLocation = ""
                            $itemSize = ""
                        }
                        
                        # Get permissions
                        $permissions = Get-ItemPermissions -ListId $list.Id -ItemId $item.Id -SiteUrl $siteUrl
                        
                        # Get sharing links
                        $sharingLinks = Get-SharingLinks -ListId $list.Id -ItemId $item.Id -SiteUrl $siteUrl
                        
                        # Prepare CSV row with proper escaping
                        $rowData = @(
                            $siteUrl.Split("/")[-1],
                            $siteUrl,
                            $list.Title,
                            $item.Id.ToString(),
                            $itemType,
                            $itemName,
                            $itemLocation,
                            $itemSize,
                            ($permissions.ReadUsers -join "`n"),
                            ($permissions.EditUsers -join "`n"),
                            ($permissions.FullControlUsers -join "`n"),
                            $sharingLinks,
                            $(if ($permissions.HasUniquePerms) { "Yes - Unique Permissions" } else { "No - Inherited from Library" }),
                            $(if ($item.Modified) { $item.Modified } else { "" }),
                            (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                        )
                        
                        # Escape each field and join with commas
                        $csvLine = ($rowData | ForEach-Object { ConvertTo-CSVCell $_ }) -join ","
                        
                        # Append to CSV file
                        Add-Content -Path $masterCsvFile -Value $csvLine -Encoding UTF8
                        $itemsWritten++
                        
                        # Show progress
                        $totalUsers = $permissions.ReadUsers.Count + $permissions.EditUsers.Count + $permissions.FullControlUsers.Count
                        $permType = if ($permissions.HasUniquePerms) { "UNIQUE" } else { "INHERITED" }
                        
                        Write-Host "  ✅ [$itemsWritten] $permType - $itemType : $itemName" -ForegroundColor $(if ($permissions.HasUniquePerms) { "Green" } else { "DarkGray" })
                        
                        # Show CSV file size every 20 items
                        if ($itemsWritten % 20 -eq 0) {
                            $fileInfo = Get-Item $masterCsvFile -ErrorAction SilentlyContinue
                            if ($fileInfo) {
                                Write-Host "  📊 CSV Status: $itemsWritten rows | Size: $([math]::Round($fileInfo.Length/1KB, 2)) KB" -ForegroundColor Cyan
                            }
                        }

                    } catch {
                        Write-Host "  ❌ Error processing item $($item.Id): $_" -ForegroundColor Red
                    }
                }
            } catch {
                Write-Host "Error retrieving items: $_" -ForegroundColor Red
                $nextPageUrl = $null
            }
        } while ($nextPageUrl)
    }

    Write-Host "`n📁 Site Summary for $($siteUrl.Split("/")[-1]):" -ForegroundColor Yellow
    Write-Host "   Items Written: $itemsWritten" -ForegroundColor White
    $fileInfo = Get-Item $masterCsvFile -ErrorAction SilentlyContinue
    if ($fileInfo) {
        Write-Host "   File Size: $([math]::Round($fileInfo.Length/1KB, 2)) KB" -ForegroundColor White
    }

    Disconnect-PnPOnline
}

Write-Host "`n====================================================================" -ForegroundColor Green
Write-Host "✅ CSV REPORT COMPLETED" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Green
Write-Host "File: $masterCsvFile" -ForegroundColor Green
$finalFileInfo = Get-Item $masterCsvFile
$totalRows = (Get-Content $masterCsvFile | Measure-Object -Line).Lines - 1
Write-Host "Total rows (including header): $($totalRows + 1)" -ForegroundColor White
Write-Host "Data rows: $totalRows" -ForegroundColor White
Write-Host "File size: $([math]::Round($finalFileInfo.Length/1KB, 2)) KB" -ForegroundColor White
Write-Host "====================================================================" -ForegroundColor Green
Write-Host "✅ FIXED ISSUES:" -ForegroundColor Yellow
Write-Host "   • Headers are now properly written at the start" -ForegroundColor White
Write-Host "   • Commas inside data are wrapped in quotes" -ForegroundColor White
Write-Host "   • Each user on new line within cell" -ForegroundColor White
Write-Host "   • Sharing links included" -ForegroundColor White
Write-Host "   • ALL items processed (both inherited and unique)" -ForegroundColor White
Write-Host "====================================================================" -ForegroundColor Green
