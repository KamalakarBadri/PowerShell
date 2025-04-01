Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Thumbprint $Thumbprint -Tenant $Tenantid
Write-Host "Connected to SharePoint site: $siteUrl" -ForegroundColor Green

# Initialize tracking
$reportData = @()
$processedItems = 0
$errorCount = 0
$startTime = Get-Date
Write-Host "Script started at $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))`n" -ForegroundColor Cyan

# Get all lists
Write-Host "Fetching all lists from site..." -ForegroundColor Yellow
$listUrl = "$siteUrl/_api/web/lists"
Write-Host "API CALL: GET $listUrl" -ForegroundColor Gray
try {
    $lists = Invoke-PnPSPRestMethod -Url $listUrl -Method Get
    Write-Host "SUCCESS: Retrieved $($lists.value.Count) lists`n" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to get lists - $_" -ForegroundColor Red
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
        Write-Host "Skipping system list: $($list.Title)" -ForegroundColor DarkGray
        continue
    }
    
    if ($list.BaseTemplate -eq 101) {
        Write-Host "`nProcessing document library: $($list.Title)" -ForegroundColor Cyan
        
        # Get all items
        $itemsUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items"
        Write-Host "API CALL: GET $itemsUrl" -ForegroundColor Gray
        try {
            $items = Invoke-PnPSPRestMethod -Url $itemsUrl -Method Get
            Write-Host "SUCCESS: Retrieved $($items.value.Count) items" -ForegroundColor Green
        } catch {
            Write-Host "ERROR: Failed to get items - $_" -ForegroundColor Red
            $errorCount++
            continue
        }

        foreach ($item in $items.value) {
            $processedItems++
            Write-Host "`nProcessing item $($item.Id) ($processedItems of $($items.value.Count))" -ForegroundColor Yellow
            
            try {
                # Get item type and details
                $itemType = if ($item.FileSystemObjectType -eq 0) { "File" } else { "Folder" }
                $itemName = $item.FieldValues.FileLeafRef ?? $item.FieldValues.Title
                Write-Host "Item Type: $itemType | Name: $itemName" -ForegroundColor White

                # Get sharing information with principal details
                $permUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation"
                Write-Host "API CALL: GET $permUrl" -ForegroundColor Gray
                $permissionsInfo = Invoke-PnPSPRestMethod -Url $permUrl -Method Get
                Write-Host "SUCCESS: Retrieved permissions information" -ForegroundColor Green

                # Process principals
                $directPermissions = @{
                    Read = @()
                    Edit = @()
                    FullControl = @()
                }

                if ($permissionsInfo.permissionsInformation.principals) {
                    Write-Host "Found $($permissionsInfo.permissionsInformation.principals.Count) principals" -ForegroundColor Cyan
                    
                    foreach ($principalElement in $permissionsInfo.permissionsInformation.principals) {
                        $principal = $principalElement.principal
                        $role = $principalElement.role
                        
                        Write-Host "`nProcessing principal:" -ForegroundColor DarkMagenta
                        Write-Host "ID: $($principal.id)" -ForegroundColor Gray
                        Write-Host "Name: $($principal.name)" -ForegroundColor Gray
                        Write-Host "LoginName: $($principal.loginName)" -ForegroundColor Gray
                        Write-Host "PrincipalType: $($principal.principalType)" -ForegroundColor Gray
                        Write-Host "Role: $role" -ForegroundColor Gray

                        if ($principal.principalType -eq 1) {
                            # User
                            $displayName = "$($principal.name) ($($principal.email ?? $principal.userPrincipalName))"
                            Write-Host "USER PERMISSION: $displayName" -ForegroundColor DarkCyan
                        }
                        elseif ($principal.principalType -in @(4,8)) {
                            # Group - get members
                            $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))/Users"
                            Write-Host "API CALL: GET $groupMembersUrl" -ForegroundColor Gray
                            
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
                                Write-Host "GROUP PERMISSION: $displayName" -ForegroundColor Magenta
                            }
                            catch {
                                $errorCount++
                                $displayName = "$($principal.name) (Failed to get members)"
                                Write-Host "ERROR: Failed to get group members - $_" -ForegroundColor Red
                            }
                        }
                        else {
                            $displayName = "$($principal.name) (Unknown type: $($principal.principalType))"
                            Write-Host "UNKNOWN PRINCIPAL TYPE: $($principal.principalType)" -ForegroundColor Yellow
                        }
                        
                        # Add to appropriate permission level
                        switch ($role) {
                            1 { 
                                $directPermissions.Read += $displayName
                                Write-Host "Permission Level: Read" -ForegroundColor DarkYellow
                            }
                            2 { 
                                $directPermissions.Edit += $displayName
                                Write-Host "Permission Level: Edit" -ForegroundColor DarkYellow
                            }
                            3 { 
                                $directPermissions.FullControl += $displayName
                                Write-Host "Permission Level: Full Control" -ForegroundColor DarkYellow
                            }
                        }
                    }
                } else {
                    Write-Host "No direct permissions (principals) found" -ForegroundColor DarkGray
                }

                # Process sharing links
                Write-Host "`nChecking sharing links..." -ForegroundColor Magenta
                $sharingLinks = @()
                if ($permissionsInfo.permissionsInformation.links) {
                    Write-Host "Found $($permissionsInfo.permissionsInformation.links.Count) sharing links" -ForegroundColor Cyan
                    
                    foreach ($link in $permissionsInfo.permissionsInformation.links) {
                        $linkDetails = $link.linkDetails
                        if ([string]::IsNullOrEmpty($linkDetails.Url)) {
                            Write-Host "Skipping link with empty URL" -ForegroundColor DarkGray
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

                        Write-Host "LINK TYPE: $linkType" -ForegroundColor White
                        Write-Host "URL: $($linkDetails.Url)" -ForegroundColor Gray
                        Write-Host "CREATED BY: $($linkDetails.CreatedBy.name)" -ForegroundColor Gray
                        if ($linkMembers) {
                            Write-Host "MEMBERS: $($linkMembers -join ', ')" -ForegroundColor Gray
                        }
                    }
                } else {
                    Write-Host "No sharing links found" -ForegroundColor DarkGray
                }

                # Add to report data
                $reportData += [PSCustomObject]@{
                    ItemID = $item.Id
                    ItemType = $itemType
                    ItemName = $itemName
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
                Write-Host "ERROR: Failed to process item $($item.Id) - $_" -ForegroundColor Red
                $reportData += [PSCustomObject]@{
                    ItemID = $item.Id
                    ItemName = "ERROR PROCESSING"
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
Write-Host "`nReport saved to: $reportFile" -ForegroundColor Green

# Display summary
Write-Host "`n=== Execution Summary ===" -ForegroundColor Cyan
Write-Host "Start Time: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))"
Write-Host "Total Items Processed: $processedItems" -ForegroundColor White
Write-Host "Total Errors Encountered: $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
Write-Host "`nScript completed!" -ForegroundColor Cyan

Disconnect-PnPOnline
