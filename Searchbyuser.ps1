## Connect to SharePoint Online
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Thumbprint $Thumbprint -Tenant $Tenantid

# Configuration
$UserUPN = "nodownload@geekbyte.online" # CASE-INSENSITIVE MATCH
$ExcludedLists = @("Access Requests", "App Packages", "appdata")

# Get all lists
$lists = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists" -Method Get
$reportData = @()

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

                        # Build report entry
                        $reportEntry = [PSCustomObject]@{
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

# Generate report
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$fileName = "$($siteUrl.Split('/')[-1])_PermissionsReport_$timestamp.csv"
$reportData | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8

Write-Host "Report generated: $fileName" -ForegroundColor Green
Disconnect-PnPOnline
