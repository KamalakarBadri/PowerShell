$TenantId = 
$ClientId = 
$ThumbPrint = 
$siteUrl = 

# Connect to SharePoint Online
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
            $hasUniqueRoleAssignments = $false
            $itemType = ""

            try {
                # Check if item has unique permissions
                $uniqueRoleAssignmentsResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/HasUniqueRoleAssignments" -Method Get
                $hasUniqueRoleAssignments = $uniqueRoleAssignmentsResponse.value

                # Only process items with unique permissions
                if ($hasUniqueRoleAssignments -eq $true) {
                    # Get item details and determine type
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
                    }

                    # Initialize permission collections
                    $fullControlUsers = @()
                    $editUsers = @()
                    $readUsers = @()
                    
                    # Get detailed permissions information for items with unique permissions
                    $permissionsInfoResponse = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation" -Method Get
                    
                    # Process direct permissions
                    if ($permissionsInfoResponse.permissionsInformation.principals) {
                        foreach ($principalInfo in $permissionsInfoResponse.permissionsInformation.principals) {
                            $principalDisplay = "$($principalInfo.principal.name) ($($principalInfo.principal.email ?? $principalInfo.principal.userPrincipalName))"
                            switch ($principalInfo.role) {
                                1 { $readUsers += $principalDisplay }
                                2 { $editUsers += $principalDisplay }
                                3 { $fullControlUsers += $principalDisplay }
                            }
                        }
                    }

                    # Process sharing links permissions
                    if ($permissionsInfoResponse.permissionsInformation.links) {
                        foreach ($link in $permissionsInfoResponse.permissionsInformation.links) {
                            # Get link members (if available)
                            if ($link.linkMembers) {
                                foreach ($member in $link.linkMembers) {
                                    $memberDisplay = "$($member.name) ($($member.email ?? $member.userPrincipalName))"
                                    
                                    # Determine permission level based on link type
                                    $linkDetails = $link.linkDetails
                                    if ($linkDetails.IsEditLink) {
                                        $editUsers += $memberDisplay
                                    } elseif ($linkDetails.IsReviewLink) {
                                        $editUsers += $memberDisplay
                                    } elseif ($linkDetails.BlocksDownload) {
                                        $readUsers += $memberDisplay	
                                    } else {
                                        $readUsers += $memberDisplay
                                    }
                                }
                            }
                        }
                    }

                    # Remove duplicate users (in case they have multiple permission sources)
                    $fullControlUsers = $fullControlUsers | Sort-Object -Unique
                    $editUsers = $editUsers | Sort-Object -Unique | Where-Object { $_ -notin $fullControlUsers }
                    $readUsers = $readUsers | Sort-Object -Unique | Where-Object { $_ -notin $fullControlUsers -and $_ -notin $editUsers }

                    # Generate report entry for items with unique permissions
                    $reportEntry = [PSCustomObject]@{
                        ItemID = $item.Id
                        ItemType = $itemType
                        Name = $itemName
                        Location = $itemLocation
                        Size = if ($itemType -eq "File") { $itemSize } else { "" }
                        
                        FullControl = $fullControlUsers -join ", "
                        Edit = $editUsers -join ", "
                        Read = $readUsers -join ", "
                        
                    }

                    # Add the entry to the report data array
                    $reportData += $reportEntry
                    
                    # Display progress
                    Write-Host "Processed: $($item.Id) [$itemType] || https://<TenantName>.sharepoint.com$itemLocation" -ForegroundColor Cyan
                }
            } catch {
                Write-Error "Error processing item $($item.Id): $_"
                # Add basic item info even if there's an error
                $reportData += [PSCustomObject]@{
                    ItemID = $item.Id
                    ItemType = "Error"
                    Name = "ERROR PROCESSING"
                    Location = ""
                    Size = ""
                   
                    FullControl = ""
                    Edit = ""
                    Read = ""
                   
                }
            }
        }
    }
}

# Export the report data to CSV
$timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
$Name = $SiteUrl.Split('/')[-1]
$reportData | Export-Csv -Path "$Name SharePoint_Unique_Permission_$timestamp.csv" -NoTypeInformation -Encoding UTF8

Write-Host "Report has been generated and saved to '$Name SharePoint_Unique_Permission_$timestamp.csv'." -ForegroundColor Green

# Disconnect from SharePoint Online
Disconnect-PnPOnline