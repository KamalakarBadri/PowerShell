$TenantId = 
$ClientId = 
$ThumbPrint = 
$siteUrl = 
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Thumbprint $Thumbprint -Tenant $Tenantid

# Retrieve all lists from the site
$lists = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists" -Method Get

# Initialize an array to store report data
$reportData = @()

# Exclude system lists
$ExcludedLists = @("Access Requests", "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", 
    "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms",
    "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", 
    "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Images", "site collection images",
    "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", 
    "Reusable Content", "Reporting Metadata", "Reporting Templates", "Search Config List", "Site Assets", 
    "Preservation Hold Library", "Site Pages", "Solution Gallery", "Style Library", "Suggested Content Browser Locations", 
    "Theme Gallery", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", 
    "Workflow History", "Workflow Tasks", "Pages")

# Iterate through each list
foreach ($list in $lists.value) {
    # Skip excluded lists
    if ($list.Title -in $ExcludedLists) { continue }
    
    # Check if the list is a document library
    if ($list.BaseTemplate -eq 101) {
        # Initialize pagination variables
        $itemsApiUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items"
        $nextPageUrl = $itemsApiUrl
        $listItems = @()

        # Loop to handle pagination
        do {
            try {
                $response = Invoke-PnPSPRestMethod -Url $nextPageUrl -Method Get
                $listItems += $response.value
                $nextPageUrl = $response."odata.nextLink" ? $response."odata.nextLink" : $null
            } catch {
                Write-Error "Failed to retrieve items from library $($list.Title). $_"
                $nextPageUrl = $null
            }
        } while ($nextPageUrl)

        # Process items from the document library
        foreach ($item in $listItems) {
            $fileSystemObjectType = $item.FileSystemObjectType
            $itemName = ""
            $itemLocation = ""
            $itemSize = 0
            $itemUniqueId = ""

            try {
                # Check if item has unique permissions (but process all items regardless)
                $uniqueRoleAssignmentsResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/HasUniqueRoleAssignments" -Method Get
                $hasUniqueRoleAssignments = $uniqueRoleAssignmentsResponse.value

                # Get item details
                if ($fileSystemObjectType -eq 0) {
                    $fileResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/file" -Method Get
                    $itemName = $fileResponse.Name
                    $itemLocation = $fileResponse.ServerRelativeUrl
                    $itemSize = $fileResponse.Length
                } elseif ($fileSystemObjectType -eq 1) {
                    $folderResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/folder" -Method Get
                    $itemName = $folderResponse.Name
                    $itemLocation = $folderResponse.ServerRelativeUrl
                }

                # Get sharing information
                $permissionsInfoResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation" -Method Get
                
                # Process direct permissions
                $directRead = @()
                $directEdit = @()
                $directFullControl = @()
                
                if ($permissionsInfoResponse.permissionsInformation.principals) {
                    foreach ($principalInfo in $permissionsInfoResponse.permissionsInformation.principals) {
                        $principalDisplay = "$($principalInfo.principal.name) ($($principalInfo.principal.email ?? $principalInfo.principal.userPrincipalName))"
                        switch ($principalInfo.role) {
                            1 { $directRead += $principalDisplay }
                            2 { $directEdit += $principalDisplay }
                            3 { $directFullControl += $principalDisplay }
                        }
                    }
                }

                # Process sharing links - filter out empty URLs
                $sharingLinks = @()
                if ($permissionsInfoResponse.permissionsInformation.links) {
                    foreach ($link in $permissionsInfoResponse.permissionsInformation.links) {
                        $linkDetails = $link.linkDetails
                        # Skip if URL is empty or null
                        if ([string]::IsNullOrEmpty($linkDetails.Url)) {
                            continue
                        }
                        
                        $linkType = switch ($true) {
                            $linkDetails.IsEditLink { "Edit" }
                            $linkDetails.IsReviewLink { "Review" }
                            default { "View" }
                        }

                        # Get link members
                        $linkMembers = @()
                        if ($link.linkMembers) {
                            foreach ($member in $link.linkMembers) {
                                $linkMembers += "$($member.name) ($($member.email ?? $member.userPrincipalName))"
                            }
                        }

                        $sharingLinks += [PSCustomObject]@{
                            LinkType = $linkType
                            Url = $linkDetails.Url
                            Created = $linkDetails.Created
                            CreatedBy = "$($linkDetails.CreatedBy.name) ($($linkDetails.CreatedBy.email ?? $linkDetails.CreatedBy.userPrincipalName))"
                            Expiration = $linkDetails.Expiration
                            IsActive = $linkDetails.IsActive
                            Members = $linkMembers -join "; "
                            ShareTokenString = $linkDetails.ShareTokenString
                            BlocksDownload = $linkDetails.BlocksDownload
                            RequiresPassword = $linkDetails.RequiresPassword
                        }
                    }
                }

                # Generate report entry with consistent columns for up to 6 links
                $reportEntry = [ordered]@{
                    ItemType = if ($fileSystemObjectType -eq 0) { "File" } else { "Folder" }
                    ItemID = $item.Id
                    ItemName = $itemName
                    ItemLocation = $itemLocation
                    ItemSize = if ($fileSystemObjectType -eq 0) { $itemSize } else { "N/A" }
                    HasUniquePermissions = $hasUniqueRoleAssignments
                    DirectRead = $directRead -join ", "
                    DirectEdit = $directEdit -join ", "
                    DirectFullControl = $directFullControl -join ", "
                    SharingLinksCount = $sharingLinks.Count
                }

                # Always include columns for up to 6 links, even if empty
                for ($i = 0; $i -lt 6; $i++) {
                    $prefix = "Link$($i+1)_"
                    if ($i -lt $sharingLinks.Count) {
                        $link = $sharingLinks[$i]
                        $reportEntry["${prefix}Type"] = $link.LinkType
                        $reportEntry["${prefix}Url"] = $link.Url
                        $reportEntry["${prefix}CreatedBy"] = $link.CreatedBy
                        $reportEntry["${prefix}Expiration"] = $link.Expiration
                        $reportEntry["${prefix}Members"] = $link.Members
                        $reportEntry["${prefix}IsActive"] = $link.IsActive
                        $reportEntry["${prefix}BlocksDownload"] = $link.BlocksDownload
                        $reportEntry["${prefix}RequiresPassword"] = $link.RequiresPassword
                    } else {
                        # Add empty columns if no link exists for this position
                        $reportEntry["${prefix}Type"] = ""
                        $reportEntry["${prefix}Url"] = ""
                        $reportEntry["${prefix}CreatedBy"] = ""
                        $reportEntry["${prefix}Expiration"] = ""
                        $reportEntry["${prefix}Members"] = ""
                        $reportEntry["${prefix}IsActive"] = ""
                        $reportEntry["${prefix}BlocksDownload"] = ""
                        $reportEntry["${prefix}RequiresPassword"] = ""
                    }
                }

                # Display the item details on screen
                Write-Host "`n--- Processing Item ---" -ForegroundColor Cyan
                Write-Host ("Item Type: {0}" -f $reportEntry.ItemType) -ForegroundColor Yellow
                Write-Host ("Item ID: {0}" -f $reportEntry.ItemID)
                Write-Host ("Item Name: {0}" -f $reportEntry.ItemName)
                Write-Host ("Item Location: {0}" -f $reportEntry.ItemLocation)
                Write-Host ("Has Unique Permissions: {0}" -f $reportEntry.HasUniquePermissions)
                if ($reportEntry.ItemSize -ne "N/A") {
                    Write-Host ("Item Size: {0} bytes" -f $reportEntry.ItemSize)
                }

                # Display direct permissions
                if ($reportEntry.DirectRead) {
                    Write-Host "`nDirect Read Access:" -ForegroundColor Green
                    $reportEntry.DirectRead -split ", " | ForEach-Object { Write-Host "- $_" }
                }
                if ($reportEntry.DirectEdit) {
                    Write-Host "`nDirect Edit Access:" -ForegroundColor Green
                    $reportEntry.DirectEdit -split ", " | ForEach-Object { Write-Host "- $_" }
                }
                if ($reportEntry.DirectFullControl) {
                    Write-Host "`nDirect Full Control Access:" -ForegroundColor Green
                    $reportEntry.DirectFullControl -split ", " | ForEach-Object { Write-Host "- $_" }
                }

                # Display sharing links
                if ($reportEntry.SharingLinksCount -gt 0) {
                    Write-Host "`nSharing Links ($($reportEntry.SharingLinksCount)):" -ForegroundColor Magenta
                    for ($i = 0; $i -lt $sharingLinks.Count; $i++) {
                        $link = $sharingLinks[$i]
                        Write-Host ("`nLink {0} ({1})" -f ($i+1), $link.LinkType) -ForegroundColor Yellow
                        Write-Host ("URL: {0}" -f $link.Url)
                        Write-Host ("Created By: {0}" -f $link.CreatedBy)
                        Write-Host ("Expires: {0}" -f $link.Expiration)
                        Write-Host ("Active: {0}" -f $link.IsActive)
                        if ($link.Members) {
                            Write-Host "Members:"
                            $link.Members -split "; " | ForEach-Object { Write-Host "- $_" }
                        }
                    }
                }

                Write-Host "`n---" -ForegroundColor Cyan

                $reportData += [PSCustomObject]$reportEntry
            } catch {
                Write-Error "Error processing item $($item.Id): $_"
                # Add basic item info even if there's an error
                $reportData += [PSCustomObject]@{
                    ItemType = if ($fileSystemObjectType -eq 0) { "File" } else { "Folder" }
                    ItemID = $item.Id
                    ItemName = "ERROR PROCESSING"
                    ItemLocation = ""
                    ItemSize = ""
                    HasUniquePermissions = "ERROR"
                    ErrorMessage = $_.Exception.Message
                }
            }
        }
    }
}

# Export the report data to CSV
$timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
$Name = $SiteUrl.Split('/')[-1]
$reportData | Export-Csv -Path "$Name SharingAccessReport_$timestamp.csv" -NoTypeInformation -Encoding UTF8

Write-Host "Report has been generated and saved to '$Name SharingAccessReport_$timestamp.csv'." -ForegroundColor Green

# Disconnect from SharePoint Online
Disconnect-PnPOnline