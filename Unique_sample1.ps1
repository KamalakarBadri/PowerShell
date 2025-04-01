$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ThumbPrint = "B799789F78628CAE56B4D0F380FD551EB754E0DB"
$siteUrl = "https://geekbyteonline.sharepoint.com/sites/New365"
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Thumbprint $Thumbprint -Tenant $Tenantid

# Initialize tracking
$reportData = @()
$processedItems = 0
$errorCount = 0
$startTime = Get-Date

# Get all lists
$listUrl = "$siteUrl/_api/web/lists"
try {
    $lists = Invoke-PnPSPRestMethod -Url $listUrl -Method Get
} catch {
    throw $_
}

# System lists to exclude
$ExcludedLists = @("Access Requests", "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", 
    "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms",
    "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", 
    "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Images", "site collection images",
    "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", 
    "Reusable Content", "Reporting Metadata", "Reporting Templates", "Search Config List", "Site Assets", 
    "Preservation Hold Library", "Site Pages", "Solution Gallery", "Style Library", "Suggested Content Browser Locations", 
    "Theme Gallery", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", 
    "Workflow History", "Workflow Tasks", "Pages")

foreach ($list in $lists.value) {
    if ($list.Title -in $ExcludedLists) {
        continue
    }
    
    if ($list.BaseTemplate -eq 101) {
        # Initialize pagination variables
        $itemsApiUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items"
        $nextPageUrl = $itemsApiUrl
        $listItems = @()

        # Pagination loop
        do {
            try {
                $response = Invoke-PnPSPRestMethod -Url $nextPageUrl -Method Get
                $listItems += $response.value

                # Check for next page
                if ($response."odata.nextLink") {
                    $nextPageUrl = $response."odata.nextLink"
                } else {
                    $nextPageUrl = $null
                }
            } catch {
                $errorCount++
                $nextPageUrl = $null
            }
        } while ($nextPageUrl)

        foreach ($item in $listItems) {
            $processedItems++

            try {
                # Check if item has unique permissions
                $uniqueRoleAssignmentsApiUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/HasUniqueRoleAssignments"
                $uniqueRoleAssignmentsResponse = Invoke-PnPSPRestMethod -Url $uniqueRoleAssignmentsApiUrl -Method Get
                $hasUniqueRoleAssignments = $uniqueRoleAssignmentsResponse.value

                # Skip items without unique permissions
                if (-not $hasUniqueRoleAssignments) {
                    continue
                }

                # Get item type and details
                $itemType = if ($item.FileSystemObjectType -eq 0) { "File" } else { "Folder" }
                $itemName = ""
                $itemLocation = ""

                if ($item.FileSystemObjectType -eq 0) {
                    # File - get file details
                    $fileUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/file"
                    try {
                        $file = Invoke-PnPSPRestMethod -Url $fileUrl -Method Get
                        $itemName = $file.Name
                        $itemLocation = $file.ServerRelativeUrl
                    } catch {
                        $itemName = $item.FieldValues.FileLeafRef ?? "Unknown"
                        $itemLocation = "Error retrieving location"
                        $errorCount++
                    }
                } else {
                    # Folder - get folder details
                    $folderUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/folder"
                    try {
                        $folder = Invoke-PnPSPRestMethod -Url $folderUrl -Method Get
                        $itemName = $folder.Name
                        $itemLocation = $folder.ServerRelativeUrl
                    } catch {
                        $itemName = $item.FieldValues.FileLeafRef ?? $item.FieldValues.Title ?? "Unknown"
                        $itemLocation = "Error retrieving location"
                        $errorCount++
                    }
                }

                # Get sharing information with principal details
                $permUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation"
                $permissionsInfo = Invoke-PnPSPRestMethod -Url $permUrl -Method Get

                # Process principals
                $directPermissions = @{
                    Read = @()
                    Edit = @()
                    FullControl = @()
                }

                if ($permissionsInfo.permissionsInformation.principals) {
                    foreach ($principalElement in $permissionsInfo.permissionsInformation.principals) {
                        $principal = $principalElement.principal
                        $role = $principalElement.role

                        if ($principal.principalType -eq 1) {
                            # User
                            $displayName = "$($principal.name) ($($principal.email ?? $principal.userPrincipalName))"
                        }
                        elseif ($principal.principalType -in @(4,8)) {
                            # Group - get members
                            $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))/Users"
                            try {
                                $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get
                                $groupMembers = @()
                                foreach ($member in $members.value) {
                                    if ($member.PrincipalType -eq 1) {
                                        $groupMembers += "$($member.Title) ($($member.UserPrincipalName ?? $member.Email))"
                                    }
                                    else {
                                        $groupMembers += "$($member.Title) (Group)"
                                    }
                                }
                                $displayName = "$($principal.name) || " + ($groupMembers -join ", ")
                            }
                            catch {
                                $errorCount++
                                $displayName = "$($principal.name) (Failed to get members)"
                            }
                        }
                        else {
                            $displayName = "$($principal.name) (Unknown type: $($principal.principalType))"
                        }
                        
                        # Add to appropriate permission level
                        switch ($role) {
                            1 { 
                                $directPermissions.Read += $displayName
                            }
                            2 { 
                                $directPermissions.Edit += $displayName
                            }
                            3 { 
                                $directPermissions.FullControl += $displayName
                            }
                        }
                    }
                }

                # Process sharing links
                $sharingLinks = @()
                if ($permissionsInfo.permissionsInformation.links) {
                    foreach ($link in $permissionsInfo.permissionsInformation.links) {
                        $linkDetails = $link.linkDetails
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
                            CreatedBy = "$($linkDetails.CreatedBy.name) ($($linkDetails.CreatedBy.email ?? $linkDetails.CreatedBy.userPrincipalName))"
                            Members = $linkMembers -join "; "
                        }
                    }
                }

                # Add to report data
                $reportData += [PSCustomObject]@{
                    ItemID = $item.Id
                    ItemType = $itemType
                    ItemName = $itemName
                    ItemLocation = $itemLocation
                    HasUniquePermissions = $hasUniqueRoleAssignments
                    ReadAccess = $directPermissions.Read -join ", "
                    EditAccess = $directPermissions.Edit -join ", "
                    FullControlAccess = $directPermissions.FullControl -join ", "
                    SharingLinksCount = $sharingLinks.Count
                    SharingLinks = ($sharingLinks | ForEach-Object { 
                        "$($_.LinkType) link: $($_.Url) (Created by: $($_.CreatedBy))" + 
                        $(if ($_.Members) { " (Members: $($_.Members))" } else { "" })
                    }) -join " | "
                }
            }
            catch {
                $errorCount++
                $reportData += [PSCustomObject]@{
                    ItemID = $item.Id
                    ItemType = "ERROR"
                    ItemName = "ERROR PROCESSING"
                    ItemLocation = "N/A"
                    Error = $_.Exception.Message
                }
            }
        }
    }
}

# Export report
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$reportFile = "CompletePermissionReport_$timestamp.csv"
$reportData | Export-Csv -Path $reportFile -NoTypeInformation -Encoding UTF8

# Display summary
Write-Host "`n=== Execution Summary ==="
Write-Host "Start Time: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))"
Write-Host "Total Items Processed: $processedItems"
Write-Host "Total Items With Unique Permissions: $($reportData.Count)"
Write-Host "Total Errors Encountered: $errorCount"

Write-Host "`nScript completed!"

Disconnect-PnPOnline
