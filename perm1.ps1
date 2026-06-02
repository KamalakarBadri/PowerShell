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

# Write CSV header initially
$csvHeader = "SiteName,SiteUrl,LibraryName,ItemID,ItemType,Name,Location,Size,ReadUsers_WithLogin,EditUsers_WithLogin,FullControlUsers_WithLogin,SharingLinks,UniquePerms,LastModified,ProcessedAt`n"
[System.IO.File]::WriteAllText($masterCsvFile, $csvHeader)

Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "DYNAMIC CSV REPORT - Updates in real-time as items are processed" -ForegroundColor Cyan
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

# Function to append a single row to CSV immediately
function Append-ToCSV {
    param($ReportEntry)
    
    $row = @(
        $ReportEntry.SiteName,
        $ReportEntry.SiteUrl,
        $ReportEntry.LibraryName,
        $ReportEntry.ItemID,
        $ReportEntry.ItemType,
        $ReportEntry.Name,
        $ReportEntry.Location,
        $ReportEntry.Size,
        $ReportEntry.ReadUsers_WithLogin,
        $ReportEntry.EditUsers_WithLogin,
        $ReportEntry.FullControlUsers_WithLogin,
        $ReportEntry.SharingLinks,
        $ReportEntry.UniquePerms,
        $ReportEntry.LastModified,
        (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    ) -join ","
    
    # Escape quotes for CSV
    $row = $row -replace '"', '""'
    
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
                            $itemName = $item.Title
                            $itemLocation = $null
                        }

                        # Check if item has unique permissions
                        try {
                            $uniquePerms = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/HasUniqueRoleAssignments" -Method Get -ErrorAction Stop
                            $hasUniquePerms = $uniquePerms.value
                        }
                        catch {
                            $hasUniquePerms = $false
                        }

                        # Initialize permission collections
                        $readUsers = @()
                        $editUsers = @()
                        $fullControlUsers = @()
                        $sharingLinks = @()

                        # Get permissions info
                        if ($hasUniquePerms) {
                            try {
                                # Get role assignments for this item
                                $roleAssignmentsUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/roleassignments?`$expand=Member,RoleDefinitionBindings"
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
                                        $userDetails = Get-UserDetails -UserId $member.Id -SiteUrl $siteUrl
                                        $userEntry = "$($userDetails.DisplayName) [$($userDetails.LoginName)]"
                                        
                                        switch ($permissionLevel) {
                                            "Read" { $readUsers += $userEntry }
                                            "Edit" { $editUsers += $userEntry }
                                            "FullControl" { $fullControlUsers += $userEntry }
                                        }
                                    }
                                    elseif ($member.PrincipalType -eq 4 -or $member.PrincipalType -eq 8) {
                                        $groupName = $member.Title
                                        $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($member.Id))/Users"
                                        try {
                                            $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                                            foreach ($groupMember in $members.value) {
                                                $userDetails = Get-UserDetails -UserId $groupMember.Id -SiteUrl $siteUrl
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
                                
                                # Check for sharing links
                                try {
                                    $sharingUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation"
                                    $sharingInfo = Invoke-PnPSPRestMethod -Url $sharingUrl -Method Get -ErrorAction Stop
                                    
                                    if ($sharingInfo.HasSharingLinks) {
                                        foreach ($link in $sharingInfo.SharingLinks) {
                                            $sharingLinks += "$($link.LinkUrl) ($($link.ShareType) access)"
                                        }
                                    }
                                }
                                catch {
                                    # No sharing links
                                }
                            }
                            catch {
                                Write-Host "    Error getting role assignments: $_" -ForegroundColor Red
                            }
                        }

                        # Create report entry
                        $reportEntry = [PSCustomObject]@{
                            SiteName = $siteUrl.Split("/")[-1]
                            SiteUrl = $siteUrl
                            LibraryName = $list.Title
                            ItemID = $item.Id
                            ItemType = $itemType
                            Name = $itemName
                            Location = $itemLocation
                            Size = if ($itemSize) { "$([math]::Round($itemSize/1KB, 2)) KB" } else { "" }
                            ReadUsers_WithLogin = ($readUsers | Sort-Object -Unique) -join "; "
                            EditUsers_WithLogin = ($editUsers | Sort-Object -Unique) -join "; "
                            FullControlUsers_WithLogin = ($fullControlUsers | Sort-Object -Unique) -join "; "
                            SharingLinks = if ($sharingLinks.Count -gt 0) { ($sharingLinks | Sort-Object -Unique) -join "; " } else { "None" }
                            UniquePerms = if ($hasUniquePerms) { "Yes" } else { "No (inherited)" }
                            LastModified = if ($item.Modified) { $item.Modified } else { "" }
                        }
                        
                        # DYNAMIC UPDATE - Append to CSV immediately
                        Append-ToCSV -ReportEntry $reportEntry
                        $itemsWithPermissions++
                        
                        # Show real-time progress
                        Write-Host "  ✅ [$($processedItems)] Added to CSV: $itemType - $itemName (Permissions: $hasUniquePerms)" -ForegroundColor Green
                        
                        # Display current CSV file size every 10 items
                        if ($itemsWithPermissions % 10 -eq 0) {
                            $fileInfo = Get-Item $masterCsvFile
                            Write-Host "  📊 CSV Status: $($itemsWithPermissions) rows written | File size: $([math]::Round($fileInfo.Length/1KB, 2)) KB" -ForegroundColor Cyan
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

    Write-Host "`n📁 Site Summary for $($siteUrl.Split("/")[-1]):" -ForegroundColor Yellow
    Write-Host "   Processed: $processedItems items" -ForegroundColor White
    Write-Host "   Written to CSV: $itemsWithPermissions items" -ForegroundColor White
    Write-Host "   CSV File Size: $([math]::Round((Get-Item $masterCsvFile).Length/1KB, 2)) KB" -ForegroundColor White

    Disconnect-PnPOnline
    Write-Host "Disconnected from $siteUrl" -ForegroundColor DarkGray
}

Write-Host "`n====================================================================" -ForegroundColor Green
Write-Host "✅ DYNAMIC CSV REPORT COMPLETED" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Green
Write-Host "Final Report: $masterCsvFile" -ForegroundColor Green
Write-Host "Total items written: $((Get-Content $masterCsvFile | Measure-Object -Line).Lines - 1)" -ForegroundColor White
Write-Host "File size: $([math]::Round((Get-Item $masterCsvFile).Length/1KB, 2)) KB" -ForegroundColor White
Write-Host "====================================================================" -ForegroundColor Green

# Optional: Open the CSV file in Excel automatically
# Invoke-Item $masterCsvFile
