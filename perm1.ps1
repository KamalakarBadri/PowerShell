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

# Create a master CSV file with timestamp
$masterTimestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$masterCsvFile = "Master_PermissionsReport_$masterTimestamp.csv"

# Function to escape CSV fields properly with newlines
function Escape-CSVField {
    param($Field)
    
    if ($Field -eq $null -or $Field -eq "") {
        return ""
    }
    
    $fieldStr = $Field.ToString()
    
    # Always wrap in double quotes and escape existing double quotes
    # This ensures newlines are preserved inside the cell
    $fieldStr = $fieldStr -replace '"', '""'
    return '"' + $fieldStr + '"'
}

# Write CSV header initially
$csvHeader = "SiteName,SiteUrl,LibraryName,ItemID,ItemType,Name,Location,Size,ReadUsers_WithLogin,EditUsers_WithLogin,FullControlUsers_WithLogin,SharingLinks,UniquePerms,LastModified,ProcessedAt`n"
[System.IO.File]::WriteAllText($masterCsvFile, $csvHeader)

Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "DYNAMIC CSV REPORT - Updates in real-time as items are processed" -ForegroundColor Cyan
Write-Host "Processing ALL items (both inherited and unique permissions)" -ForegroundColor Yellow
Write-Host "Including Sharing Links" -ForegroundColor Yellow
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
        # Method 1: Try to get sharing information
        $sharingUrl = "$SiteUrl/_api/web/lists(guid'$ListId')/items($ItemId)/GetSharingInformation"
        $sharingInfo = Invoke-PnPSPRestMethod -Url $sharingUrl -Method Get -ErrorAction Stop
        
        if ($sharingInfo.HasSharingLinks -and $sharingInfo.SharingLinks) {
            foreach ($link in $sharingInfo.SharingLinks) {
                $linkInfo = "$($link.LinkUrl) [$($link.ShareType) access"
                if ($link.ExpirationDate) {
                    $linkInfo += " - Expires: $($link.ExpirationDate)"
                }
                if ($link.PasswordProtected) {
                    $linkInfo += " - Password Protected"
                }
                $linkInfo += "]"
                $sharingLinks += $linkInfo
            }
        }
        
        # Method 2: Try to get role assignments that might be sharing links
        $roleAssignmentsUrl = "$SiteUrl/_api/web/lists(guid'$ListId')/items($ItemId)/roleassignments"
        $roleAssignments = Invoke-PnPSPRestMethod -Url $roleAssignmentsUrl -Method Get -ErrorAction SilentlyContinue
        
        if ($roleAssignments -and $roleAssignments.value) {
            foreach ($ra in $roleAssignments.value) {
                if ($ra.Member.PrincipalType -eq 4 -and $ra.Member.Title -like "*SharingLink*") {
                    $sharingLinks += "Sharing Link: $($ra.Member.Title)"
                }
            }
        }
    }
    catch {
        # No sharing links found or error
    }
    
    return $sharingLinks
}

# Function to get permissions for an item (both direct and inherited)
function Get-ItemPermissions {
    param($ListId, $ItemId, $SiteUrl)
    
    $readUsers = @()
    $editUsers = @()
    $fullControlUsers = @()
    
    try {
        # First, check if item has unique permissions
        $uniquePermsCheck = Invoke-PnPSPRestMethod -Url "$SiteUrl/_api/web/lists(guid'$ListId')/items($ItemId)/HasUniqueRoleAssignments" -Method Get -ErrorAction Stop
        $hasUniquePerms = $uniquePermsCheck.value
        
        if ($hasUniquePerms) {
            # Get role assignments for this specific item
            $roleAssignmentsUrl = "$SiteUrl/_api/web/lists(guid'$ListId')/items($ItemId)/roleassignments?`$expand=Member,RoleDefinitionBindings"
            $roleAssignments = Invoke-PnPSPRestMethod -Url $roleAssignmentsUrl -Method Get -ErrorAction Stop
            
            foreach ($ra in $roleAssignments.value) {
                $member = $ra.Member
                $roleBindings = $ra.RoleDefinitionBindings
                
                # Determine permission level
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
                    # User
                    $userDetails = Get-UserDetails -UserId $member.Id -SiteUrl $SiteUrl
                    $userEntry = "$($userDetails.DisplayName) [$($userDetails.LoginName)]"
                    
                    switch ($permissionLevel) {
                        "Read" { $readUsers += $userEntry }
                        "Edit" { $editUsers += $userEntry }
                        "FullControl" { $fullControlUsers += $userEntry }
                    }
                }
                elseif ($member.PrincipalType -eq 4 -or $member.PrincipalType -eq 8) {
                    # SharePoint Group
                    $groupName = $member.Title
                    $groupMembersUrl = "$SiteUrl/_api/web/SiteGroups/GetById($($member.Id))/Users"
                    try {
                        $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                        foreach ($groupMember in $members.value) {
                            $userDetails = Get-UserDetails -UserId $groupMember.Id -SiteUrl $SiteUrl
                            $userEntry = "$($userDetails.DisplayName) [$($userDetails.LoginName)] (via $groupName)"
                            
                            switch ($permissionLevel) {
                                "Read" { $readUsers += $userEntry }
                                "Edit" { $editUsers += $userEntry }
                                "FullControl" { $fullControlUsers += $userEntry }
                            }
                        }
                    }
                    catch {
                        $groupEntry = "$groupName [Group - members not accessible]"
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
            # Item inherits permissions - get from list/parent
            $listRoleAssignmentsUrl = "$SiteUrl/_api/web/lists(guid'$ListId')/roleassignments?`$expand=Member,RoleDefinitionBindings"
            $listRoleAssignments = Invoke-PnPSPRestMethod -Url $listRoleAssignmentsUrl -Method Get -ErrorAction Stop
            
            foreach ($ra in $listRoleAssignments.value) {
                $member = $ra.Member
                $roleBindings = $ra.RoleDefinitionBindings
                
                # Determine permission level
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
                    # User
                    $userDetails = Get-UserDetails -UserId $member.Id -SiteUrl $SiteUrl
                    $userEntry = "$($userDetails.DisplayName) [$($userDetails.LoginName)] (inherited from library)"
                    
                    switch ($permissionLevel) {
                        "Read" { $readUsers += $userEntry }
                        "Edit" { $editUsers += $userEntry }
                        "FullControl" { $fullControlUsers += $userEntry }
                    }
                }
                elseif ($member.PrincipalType -eq 4 -or $member.PrincipalType -eq 8) {
                    # SharePoint Group
                    $groupName = $member.Title
                    $groupMembersUrl = "$SiteUrl/_api/web/SiteGroups/GetById($($member.Id))/Users"
                    try {
                        $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                        foreach ($groupMember in $members.value) {
                            $userDetails = Get-UserDetails -UserId $groupMember.Id -SiteUrl $SiteUrl
                            $userEntry = "$($userDetails.DisplayName) [$($userDetails.LoginName)] (via $groupName - inherited from library)"
                            
                            switch ($permissionLevel) {
                                "Read" { $readUsers += $userEntry }
                                "Edit" { $editUsers += $userEntry }
                                "FullControl" { $fullControlUsers += $userEntry }
                            }
                        }
                    }
                    catch {
                        $groupEntry = "$groupName [Group - members not accessible] (inherited from library)"
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
    
    return @{
        ReadUsers = $readUsers | Sort-Object -Unique
        EditUsers = $editUsers | Sort-Object -Unique
        FullControlUsers = $fullControlUsers | Sort-Object -Unique
        HasUniquePerms = $hasUniquePerms
    }
}

# Function to append a single row to CSV immediately with proper escaping
function Append-ToCSV {
    param($ReportEntry)
    
    # Escape each field (newlines will be preserved inside quotes)
    $fields = @(
        (Escape-CSVField $ReportEntry.SiteName),
        (Escape-CSVField $ReportEntry.SiteUrl),
        (Escape-CSVField $ReportEntry.LibraryName),
        (Escape-CSVField $ReportEntry.ItemID),
        (Escape-CSVField $ReportEntry.ItemType),
        (Escape-CSVField $ReportEntry.Name),
        (Escape-CSVField $ReportEntry.Location),
        (Escape-CSVField $ReportEntry.Size),
        (Escape-CSVField $ReportEntry.ReadUsers_WithLogin),
        (Escape-CSVField $ReportEntry.EditUsers_WithLogin),
        (Escape-CSVField $ReportEntry.FullControlUsers_WithLogin),
        (Escape-CSVField $ReportEntry.SharingLinks),
        (Escape-CSVField $ReportEntry.UniquePerms),
        (Escape-CSVField $ReportEntry.LastModified),
        (Escape-CSVField (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
    )
    
    $row = $fields -join ","
    
    # Add the row to CSV file
    Add-Content -Path $masterCsvFile -Value $row -Encoding UTF8
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
    $itemsWithPermissions = 0

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
                        # Get item details and determine type
                        $itemName = $null
                        $itemLocation = $null
                        $itemSize = $null
                        $itemType = "Unknown"
                        
                        if ($item.FileSystemObjectType -eq 0 -and $item.File) {
                            $itemType = "File"
                            $itemName = $item.File.Name
                            $itemLocation = $item.File.ServerRelativeUrl
                            $itemSize = $item.File.Length
                        } 
                        elseif ($item.FileSystemObjectType -eq 1 -and $item.Folder) {
                            $itemType = "Folder"
                            $itemName = $item.Folder.Name
                            $itemLocation = $item.Folder.ServerRelativeUrl
                        } 
                        else {
                            $itemType = "ListItem"
                            $itemName = if ($item.Title) { $item.Title } else { "Item_$($item.Id)" }
                            $itemLocation = $null
                        }
                        
                        # Get permissions for this item (handles both inherited and unique)
                        $permissions = Get-ItemPermissions -ListId $list.Id -ItemId $item.Id -SiteUrl $siteUrl
                        
                        # Get sharing links for this item
                        $sharingLinks = Get-SharingLinks -ListId $list.Id -ItemId $item.Id -SiteUrl $siteUrl
                        $sharingLinksText = if ($sharingLinks.Count -gt 0) { $sharingLinks -join "`n" } else { "None" }
                        
                        # Create report entry with newlines
                        $readUsersText = if ($permissions.ReadUsers.Count -gt 0) { $permissions.ReadUsers -join "`n" } else { "No direct read access" }
                        $editUsersText = if ($permissions.EditUsers.Count -gt 0) { $permissions.EditUsers -join "`n" } else { "No direct edit access" }
                        $fullControlUsersText = if ($permissions.FullControlUsers.Count -gt 0) { $permissions.FullControlUsers -join "`n" } else { "No full control access" }
                        
                        $reportEntry = [PSCustomObject]@{
                            SiteName = $siteUrl.Split("/")[-1]
                            SiteUrl = $siteUrl
                            LibraryName = $list.Title
                            ItemID = $item.Id
                            ItemType = $itemType
                            Name = $itemName
                            Location = $itemLocation
                            Size = if ($itemSize) { "$([math]::Round($itemSize/1KB, 2)) KB" } else { "" }
                            ReadUsers_WithLogin = $readUsersText
                            EditUsers_WithLogin = $editUsersText
                            FullControlUsers_WithLogin = $fullControlUsersText
                            SharingLinks = $sharingLinksText
                            UniquePerms = if ($permissions.HasUniquePerms) { "Yes" } else { "No (inherited from library)" }
                            LastModified = if ($item.Modified) { $item.Modified } else { "" }
                        }
                        
                        # DYNAMIC UPDATE - Append to CSV immediately
                        Append-ToCSV -ReportEntry $reportEntry
                        $itemsWithPermissions++
                        
                        # Show real-time progress
                        $totalUsers = $permissions.ReadUsers.Count + $permissions.EditUsers.Count + $permissions.FullControlUsers.Count
                        $permType = if ($permissions.HasUniquePerms) { "UNIQUE" } else { "INHERITED" }
                        $sharingCount = $sharingLinks.Count
                        
                        Write-Host "  📝 [$($processedItems)] $permType - $itemType : $itemName (Users: $totalUsers, Sharing Links: $sharingCount)" -ForegroundColor $(if ($permissions.HasUniquePerms) { "Green" } else { "DarkGray" })
                        
                        # Display current CSV file size every 20 items
                        if ($itemsWithPermissions % 20 -eq 0) {
                            $fileInfo = Get-Item $masterCsvFile -ErrorAction SilentlyContinue
                            if ($fileInfo) {
                                Write-Host "  📊 CSV Status: $($itemsWithPermissions) rows written | File size: $([math]::Round($fileInfo.Length/1KB, 2)) KB" -ForegroundColor Cyan
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
    Write-Host "   Total items processed: $processedItems" -ForegroundColor White
    Write-Host "   Items written to CSV: $itemsWithPermissions" -ForegroundColor White
    $fileInfo = Get-Item $masterCsvFile -ErrorAction SilentlyContinue
    if ($fileInfo) {
        Write-Host "   CSV File Size: $([math]::Round($fileInfo.Length/1KB, 2)) KB" -ForegroundColor White
    }

    Disconnect-PnPOnline
    Write-Host "Disconnected from $siteUrl" -ForegroundColor DarkGray
}

Write-Host "`n====================================================================" -ForegroundColor Green
Write-Host "✅ DYNAMIC CSV REPORT COMPLETED" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Green
Write-Host "Final Report: $masterCsvFile" -ForegroundColor Green
$finalFileInfo = Get-Item $masterCsvFile
$totalRows = (Get-Content $masterCsvFile | Measure-Object -Line).Lines - 1
Write-Host "Total items written: $totalRows" -ForegroundColor White
Write-Host "File size: $([math]::Round($finalFileInfo.Length/1KB, 2)) KB" -ForegroundColor White
Write-Host "====================================================================" -ForegroundColor Green
Write-Host "✅ INCLUDED FOR EACH ITEM:" -ForegroundColor Yellow
Write-Host "   • ALL items (both inherited and unique permissions)" -ForegroundColor White
Write-Host "   • Sharing links with expiration and password details" -ForegroundColor White
Write-Host "   • User login names with display names" -ForegroundColor White
Write-Host "   • Group memberships expanded to individual users" -ForegroundColor White
Write-Host "   • Each user on new line within cell (enable Wrap Text in Excel)" -ForegroundColor White
Write-Host "====================================================================" -ForegroundColor Green

# Optional: Open the CSV file in Excel automatically
# Invoke-Item $masterCsvFile
