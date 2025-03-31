
# List of SharePoint Sites
$SiteUrls = @(
   
)

# List of Users to Check Permissions
$UserUPNs = @(
    
)

# Loop through each SharePoint site
foreach ($siteUrl in $SiteUrls) {
    Write-Host "Connecting to: $siteUrl" -ForegroundColor Cyan
    Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId

    # Get all document libraries
    $lists = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists" -Method Get
    $reportData = @()

    foreach ($list in $lists.value) {
        if ($list.BaseTemplate -ne 101) { continue } # Process only document libraries
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

                        # Determine if the item is a file or folder
                        $itemType = switch ($item.FileSystemObjectType) {
                            0 { "File" }
                            1 { "Folder" }
                            default { "ListItem" }
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
                                    ItemID      = $item.Id
                                    UserUPN     = $UserUPN
                                    ItemType    = $itemType
                                    Name        = $item.Title
                                    Location    = $item.ServerRelativeUrl
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

    # Generate CSV report
    if ($reportData.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $siteName = $siteUrl.Split("/")[-1]
        $fileName = "$siteName-FilteredPermissionsReport_$timestamp.csv"
        $reportData | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8

        Write-Host "Filtered report generated: $fileName" -ForegroundColor Green
    } else {
        Write-Host "No matching permissions found for users in $siteUrl." -ForegroundColor Yellow
    }

    Disconnect-PnPOnline
}
