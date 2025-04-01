$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ThumbPrint = "B799789F78628CAE56B4D0F380FD551EB754E0DB"

# Array of site URLs to process
$siteUrls = @(
    "https://geekbyteonline.sharepoint.com/sites/New365",
    "https://geekbyteonline.sharepoint.com/sites/AnotherSite",
    "https://geekbyteonline.sharepoint.com/sites/ThirdSite"
)

# Array of users to check (case-insensitive)
$usersToCheck = @(
    "nodownload@geekbyte.online",
    "anotheruser@geekbyte.online",
    "thirduser@geekbyte.online"
)

# Exclude system lists
$ExcludedLists = @("Access Requests", "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", 
    "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms",
    "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", 
    "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Images", "site collection images",
    "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", 
    "Reusable Content", "Reporting Metadata", "Reporting Templates", "Search Config List", "Site Assets", 
    "Preservation Hold Library", "Solution Gallery", "Style Library", "Suggested Content Browser Locations", 
    "Theme Gallery", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", 
    "Workflow History", "Workflow Tasks", "Pages")

$reportData = @()

foreach ($siteUrl in $siteUrls) {
    try {
        # Connect to SharePoint Online
        Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Thumbprint $Thumbprint -Tenant $Tenantid -ErrorAction Stop
        
        Write-Host "Connected to site: $siteUrl" -ForegroundColor Green

        # Get all lists
        $lists = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists" -Method Get
        
        foreach ($list in $lists.value) {
            if ($list.Title -in $ExcludedLists) { continue }
            
            if ($list.BaseTemplate -eq 101) {
                Write-Host "Processing Document Library: $($list.Title)" -ForegroundColor Cyan
                
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

                                # Get item details
                                $itemType = switch ($item.FileSystemObjectType) {
                                    0 { 
                                        $file = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/file" -Method Get
                                        "File"
                                    }
                                    1 { 
                                        $folder = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/folder" -Method Get
                                        "Folder"
                                    }
                                    default { "ListItem" }
                                }

                                # Get sharing information with principal details
                                $permUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation"
                                $permissionsInfo = Invoke-PnPSPRestMethod -Url $permUrl -Method Get

                                # Check each user for permissions
                                foreach ($UserUPN in $usersToCheck) {
                                    $hasAccess = $false
                                    $readSources = @()
                                    $editSources = @()
                                    $fullControl = $false
                                    $sharingLinks = @()

                                    # Process direct permissions and group memberships
                                    if ($permissionsInfo.permissionsInformation.principals) {
                                        foreach ($principalElement in $permissionsInfo.permissionsInformation.principals) {
                                            $principal = $principalElement.principal
                                            $role = $principalElement.role

                                            if ($principal.principalType -eq 1) {
                                                # User - check if it's our target user
                                                $principalUpn = $principal.userPrincipalName ?? $principal.email
                                                if ($principalUpn -like "*$UserUPN*") {
                                                    $hasAccess = $true
                                                    switch ($role) {
                                                        1 { $readSources += "Direct Permission" }
                                                        2 { $editSources += "Direct Permission" }
                                                        3 { $fullControl = $true }
                                                    }
                                                }
                                            }
                                            elseif ($principal.principalType -in @(4,8)) {
                                                # Group - get members and check if our user is a member
                                                $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($principal.id))/Users"
                                                try {
                                                    $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get
                                                    foreach ($member in $members.value) {
                                                        if ($member.PrincipalType -eq 1) {
                                                            $memberUpn = $member.UserPrincipalName ?? $member.Email
                                                            if ($memberUpn -like "*$UserUPN*") {
                                                                $hasAccess = $true
                                                                switch ($role) {
                                                                    1 { $readSources += "Member of '$($principal.name)' group" }
                                                                    2 { $editSources += "Member of '$($principal.name)' group" }
                                                                    3 { $fullControl = $true }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                catch {
                                                    Write-Host "Error getting members for group $($principal.name): $_" -ForegroundColor Yellow
                                                }
                                            }
                                        }
                                    }

                                    # Process sharing links
                                    if ($permissionsInfo.permissionsInformation.links) {
                                        foreach ($link in $permissionsInfo.permissionsInformation.links) {
                                            $linkDetails = $link.linkDetails
                                            $linkUrl = $linkDetails.Url
                                            
                                            if ($link.linkMembers) {
                                                foreach ($member in $link.linkMembers) {
                                                    $memberUpn = $member.userPrincipalName ?? $member.email
                                                    
                                                    if ($memberUpn -like "*$UserUPN*") {
                                                        $hasAccess = $true
                                                        $linkType = switch ($linkDetails.LinkKind) {
                                                            0 { "View" }
                                                            1 { "Edit" }
                                                            2 { "Review" }
                                                            3 { "BlockedDownload" }
                                                            4 { "OrganizationView" }
                                                            5 { "OrganizationEdit" }
                                                            default { "Unknown" }
                                                        }
                                                        
                                                        $sharingLinks += "$linkType link: $linkUrl"
                                                        
                                                        if ($linkDetails.IsEditLink -or $linkDetails.IsReviewLink) {
                                                            $editSources += "Sharing link ($linkUrl)"
                                                        }
                                                        else {
                                                            $readSources += "Sharing link ($linkUrl)"
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    # Only add to report if user has any permissions
                                    if ($hasAccess) {
                                        # Build report entry
                                        $reportEntry = [PSCustomObject]@{
                                            Site        = $siteUrl
                                            User        = $UserUPN
                                            ItemID      = $item.Id
                                            ItemType    = $itemType
                                            Name        = if ($itemType -eq "File") { $file.Name } 
                                                        elseif ($itemType -eq "Folder") { $folder.Name } 
                                                        else { $item.Title }
                                            Location    = if ($itemType -eq "File") { $file.ServerRelativeUrl } 
                                                        elseif ($itemType -eq "Folder") { $folder.ServerRelativeUrl } 
                                                        else { "" }
                                            Size        = if ($itemType -eq "File") { $file.Length } else { "" }
                                            Read        = if ($readSources.Count -gt 0) { $readSources -join ", " } else { "" }
                                            Edit        = if ($editSources.Count -gt 0) { $editSources -join ", " } else { "" }
                                            FullControl = if ($fullControl) { "Yes" } else { "" }
                                            SharingLinks = if ($sharingLinks.Count -gt 0) { $sharingLinks -join "`n" } else { "" }
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
        }
    } catch {
        Write-Host "Error connecting to site $siteUrl : $_" -ForegroundColor Red
    } finally {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
}

# Generate report only if we have data
if ($reportData.Count -gt 0) {
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $fileName = "UserAccessReport_$timestamp.csv"
    $reportData | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8
    Write-Host "Report generated: $fileName" -ForegroundColor Green
} else {
    Write-Host "No items with user access found." -ForegroundColor Yellow
}
