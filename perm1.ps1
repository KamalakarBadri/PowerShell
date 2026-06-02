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
$masterCsvFile = "SharePoint_Permissions_Report_$masterTimestamp.csv"

# Simple function to escape CSV fields
function Write-CSVRow {
    param($Values)
    
    $escapedValues = @()
    foreach ($value in $Values) {
        if ($value -eq $null) { 
            $escapedValues += ""
            continue
        }
        
        $strValue = $value.ToString()
        
        # If contains comma, newline, or double quote, wrap in quotes and escape quotes
        if ($strValue -match '["\n\r,]') {
            $strValue = $strValue -replace '"', '""'
            $escapedValues += "`"$strValue`""
        }
        else {
            $escapedValues += $strValue
        }
    }
    
    return $escapedValues -join ","
}

# Write header
$headers = "SiteName,SiteUrl,LibraryName,ItemID,ItemType,Name,Location,SizeKB,ReadUsers,EditUsers,FullControlUsers,SharingLinks,UniquePerms,LastModified,ProcessedAt"
[System.IO.File]::WriteAllText($masterCsvFile, $headers + "`n")

Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "SHAREPOINT PERMISSIONS REPORT" -ForegroundColor Cyan
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
                            $itemSize = [math]::Round($item.File.Length/1KB, 2)
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

                        # Check unique permissions
                        $hasUniquePerms = $false
                        try {
                            $uniquePerms = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/HasUniqueRoleAssignments" -Method Get -ErrorAction Stop
                            $hasUniquePerms = $uniquePerms.value
                        }
                        catch {
                            $hasUniquePerms = $false
                        }

                        # Initialize arrays
                        $readUsers = @()
                        $editUsers = @()
                        $fullControlUsers = @()
                        $sharingLinks = @()

                        # Get permissions
                        try {
                            $permsInfo = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation" -Method Get -ErrorAction Stop
                            
                            if ($permsInfo.permissionsInformation.principals) {
                                foreach ($principalElement in $permsInfo.permissionsInformation.principals) {
                                    $principal = $principalElement.principal
                                    $role = $principalElement.role

                                    if ($principal.principalType -eq 1) {
                                        # User
                                        $userLogin = if ($principal.userPrincipalName) { $principal.userPrincipalName } else { $principal.email }
                                        $userName = if ($principal.name) { $principal.name } else { $userLogin }
                                        $userText = "$userName [$userLogin]"
                                        
                                        if (-not $hasUniquePerms) {
                                            $userText = "$userText [INHERITED]"
                                        }
                                        
                                        switch ($role) {
                                            1 { $readUsers += $userText }
                                            2 { $editUsers += $userText }
                                            3 { $fullControlUsers += $userText }
                                        }
                                    }
                                    elseif ($principal.principalType -in @(4,8)) {
                                        # Group
                                        $groupName = $principal.name
                                        $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))/Users"
                                        try {
                                            $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                                            foreach ($member in $members.value) {
                                                if ($member.PrincipalType -eq 1) {
                                                    $memberLogin = if ($member.UserPrincipalName) { $member.UserPrincipalName } else { $member.Email }
                                                    $memberName = if ($member.Title) { $member.Title } else { $memberLogin }
                                                    $memberText = "$memberName [$memberLogin] (via $groupName)"
                                                    
                                                    if (-not $hasUniquePerms) {
                                                        $memberText = "$memberText [INHERITED]"
                                                    }
                                                    
                                                    switch ($role) {
                                                        1 { $readUsers += $memberText }
                                                        2 { $editUsers += $memberText }
                                                        3 { $fullControlUsers += $memberText }
                                                    }
                                                }
                                            }
                                        }
                                        catch {
                                            $groupText = "$groupName [GROUP MEMBERS NOT ACCESSIBLE]"
                                            if (-not $hasUniquePerms) {
                                                $groupText = "$groupText [INHERITED]"
                                            }
                                            switch ($role) {
                                                1 { $readUsers += $groupText }
                                                2 { $editUsers += $groupText }
                                                3 { $fullControlUsers += $groupText }
                                            }
                                        }
                                    }
                                }
                            }

                            # Get sharing links
                            if ($permsInfo.permissionsInformation.links) {
                                foreach ($link in $permsInfo.permissionsInformation.links) {
                                    $linkUrl = $link.linkDetails.Url
                                    $linkType = if ($link.linkDetails.IsEditLink) { "Edit" } else { "Read" }
                                    $linkText = "$linkUrl ($linkType access)"
                                    
                                    if ($link.linkDetails.ExpirationDate) {
                                        $linkText = "$linkText - Expires: $($link.linkDetails.ExpirationDate)"
                                    }
                                    
                                    $sharingLinks += $linkText
                                }
                            }
                        }
                        catch {
                            # No permissions found
                        }

                        # Default values if no data
                        if ($readUsers.Count -eq 0 -and $editUsers.Count -eq 0 -and $fullControlUsers.Count -eq 0) {
                            $readUsers = @("No permissions found")
                        }
                        
                        if ($sharingLinks.Count -eq 0) {
                            $sharingLinks = @("No sharing links")
                        }

                        # Join with semicolon and space (not newline) to avoid CSV complexity
                        $readText = $readUsers -join "; "
                        $editText = $editUsers -join "; "
                        $fullText = $fullControlUsers -join "; "
                        $linksText = $sharingLinks -join "; "

                        # Prepare row data
                        $rowValues = @(
                            $siteUrl.Split("/")[-1],
                            $siteUrl,
                            $list.Title,
                            $item.Id,
                            $itemType,
                            $itemName,
                            $itemLocation,
                            $itemSize,
                            $readText,
                            $editText,
                            $fullText,
                            $linksText,
                            $(if ($hasUniquePerms) { "Yes" } else { "No (Inherited)" }),
                            $(if ($item.Modified) { $item.Modified } else { "" }),
                            (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                        )

                        # Write to CSV
                        $csvRow = Write-CSVRow -Values $rowValues
                        Add-Content -Path $masterCsvFile -Value $csvRow -Encoding UTF8
                        $itemsWritten++

                        Write-Host "      ✅ Added: $itemType - $itemName" -ForegroundColor Green

                    } catch {
                        Write-Host "      ❌ Error: $_" -ForegroundColor Red
                    }
                }
            } catch {
                Write-Host "  Error retrieving items: $_" -ForegroundColor Red
                $nextPageUrl = $null
            }
        } while ($nextPageUrl)
    }

    Write-Host "`n📁 Site Summary: $($itemsWritten) items written to CSV" -ForegroundColor Yellow
    Disconnect-PnPOnline
}

Write-Host "`n====================================================================" -ForegroundColor Green
Write-Host "✅ REPORT COMPLETED" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Green
Write-Host "File: $masterCsvFile" -ForegroundColor Green

$fileInfo = Get-Item $masterCsvFile
$lines = Get-Content $masterCsvFile
Write-Host "Total rows: $($lines.Count)" -ForegroundColor White
Write-Host "File size: $([math]::Round($fileInfo.Length/1KB, 2)) KB" -ForegroundColor White
Write-Host "====================================================================" -ForegroundColor Green

# Show first few lines to verify
Write-Host "`nFirst 3 lines of CSV for verification:" -ForegroundColor Yellow
Get-Content $masterCsvFile -TotalCount 3 | ForEach-Object { Write-Host $_ -ForegroundColor Cyan }
Write-Host "====================================================================" -ForegroundColor Green
