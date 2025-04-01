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
    "Contribute@geekbyte.online"
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
    Write-Host "Connecting to: $siteUrl" -ForegroundColor Cyan
    Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId

    # Get all document libraries
    $lists = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists" -Method Get
    $reportData = @()

    foreach ($list in $lists.value) {
        if ($list.BaseTemplate -ne 101) { continue } # Process only document libraries

        # Check if the list is in the exclusion list
        if ($ExcludedLists -contains $list.Title) {
            Write-Host "Skipping excluded list: $($list.Title)" -ForegroundColor Gray
            continue
        }

        Write-Host "Processing Document Library: $($list.Title)" -ForegroundColor Yellow

        $nextPageUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items?`$top=1000"
        do {
            try {
                $response = Invoke-PnPSPRestMethod -Url $nextPageUrl -Method Get
                $listItems = $response.value
                $nextPageUrl = $response."odata.nextLink"

                foreach ($item in $listItems) {
                    try {
                        # Check unique permissions
                        $uniquePerms = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/HasUniqueRoleAssignments" -Method Get
                        if (-not $uniquePerms.value) { continue }

                        # Get item details and determine type
                        $fileSystemObjectType = $item.FileSystemObjectType
                        $itemName = $null
                        $itemLocation = $null
                        $itemSize = $null
                        
                        if ($fileSystemObjectType -eq 0) {
                            $itemType = "File"
                            $fileResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/file" -Method Get
                            $itemName = $fileResponse.Name
                            $itemLocation = $fileResponse.ServerRelativeUrl
                            $itemSize = $fileResponse.Length
                        } elseif ($fileSystemObjectType -eq 1) {
                            $itemType = "Folder"
                            $folderResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/folder" -Method Get
                            $itemName = $folderResponse.Name
                            $itemLocation = $folderResponse.ServerRelativeUrl
                        } else {
                            $itemType = "ListItem"
                            $itemName = $item.Title
                            $itemLocation = $null
                        }

                        # Get permissions info
                        $permsInfo = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation" -Method Get

                        foreach ($UserUPN in $UserUPNs) {
                            $readSources = @()
                            $editSources = @()
                            $fullControl = $false

                            # Process direct permissions
                            if ($permsInfo.permissionsInformation.principals) {
                                foreach ($principal in $permsInfo.permissionsInformation.principals) {
                                    $principalUpn = $principal.principal.userPrincipalName ?? $principal.principal.email
                                    if ($principalUpn -like "*$UserUPN*") { 
                                        switch ($principal.role) {
                                            1 { $readSources += "Direct Permission" }
                                            2 { $editSources += "Direct Permission" }
                                            3 { $fullControl = $true }
                                        }
                                    }
                                }
                            }

                            # Process sharing links
                            if ($permsInfo.permissionsInformation.links) {
                                foreach ($link in $permsInfo.permissionsInformation.links) {
                                    $linkUrl = $link.linkDetails.Url
                                    
                                    if ($link.linkMembers) {
                                        foreach ($member in $link.linkMembers) {
                                            $memberUpn = $member.userPrincipalName ?? $member.email
                                            if ($memberUpn -like "*$UserUPN*") {
                                                if ($link.linkDetails.IsEditLink -or $link.linkDetails.IsReviewLink) {
                                                    $editSources += $linkUrl
                                                }
                                                else {
                                                    $readSources += $linkUrl
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            # Only add to report if the user has Read, Edit, or Full Control
                            if ($readSources -or $editSources -or $fullControl) {
                                $reportEntry = [PSCustomObject]@{
                                    SiteName    = $siteUrl.Split("/")[-1]
                                    LibraryName = $list.Title
                                    ItemID      = $item.Id
                                    UserUPN     = $UserUPN
                                    ItemType    = $itemType
                                    Name        = $itemName
                                    Location    = $itemLocation
                                    Size        = if ($itemSize) { "$([math]::Round($itemSize/1KB, 2)) KB" } else { "" }
                                    Read        = if ($readSources.Count -gt 0) { $readSources -join "`n" } else { "" }
                                    Edit        = if ($editSources.Count -gt 0) { $editSources -join "`n" } else { "" }
                                    FullControl = if ($fullControl) { "Yes" } else { "" }
                                }
                                $reportData += $reportEntry
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
        $reportData | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8

        Write-Host "Report generated: $fileName" -ForegroundColor Green
    } else {
        Write-Host "No matching permissions found for users in $siteUrl." -ForegroundColor Yellow
    }

    Disconnect-PnPOnline
}
