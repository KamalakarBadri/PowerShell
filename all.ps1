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

                                # Get permissions info
                                $permsInfo = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation?`$expand=permissionsInformation" -Method Get

                                # Check each user for permissions
                                foreach ($UserUPN in $usersToCheck) {
                                    $readSources = @()
                                    $editSources = @()
                                    $fullControl = $false

                                    # Process direct permissions
                                    if ($permsInfo.permissionsInformation.principals) {
                                        foreach ($principal in $permsInfo.permissionsInformation.principals) {
                                            $principalUpn = $principal.principal.userPrincipalName ?? $principal.principal.email
                                            if ($principalUpn -like "*$UserUPN*") { # Wildcard match for case insensitivity
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
                                            Write-Host "Processing link: $linkUrl" -ForegroundColor DarkGray
                                            
                                            if ($link.linkMembers) {
                                                foreach ($member in $link.linkMembers) {
                                                    $memberUpn = $member.userPrincipalName ?? $member.email
                                                    Write-Host "Checking member: $memberUpn" -ForegroundColor DarkGray
                                                    
                                                    # Case-insensitive comparison
                                                    if ($memberUpn -like "*$UserUPN*") {
                                                        Write-Host "Match found for $UserUPN in link: $linkUrl" -ForegroundColor Green
                                                        
                                                        # Determine permission type
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

                                    # Only add to report if user has any permissions
                                    if ($readSources.Count -gt 0 -or $editSources.Count -gt 0 -or $fullControl) {
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
                                            Read        = $readSources -join "`n"
                                            Edit        = $editSources -join "`n"
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
        }
    } catch {
        Write-Host "Error connecting to site $siteUrl : $_" -ForegroundColor Red
    } finally {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
}

# Generate report
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$fileName = "Combined_PermissionsReport_$timestamp.csv"
$reportData | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8

Write-Host "Report generated: $fileName" -ForegroundColor Green
