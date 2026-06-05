$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ThumbPrint = "B799789F78628CAE56B4D0F380FD551EB754E0DB"

# Array of site URLs to process
$SiteUrls = @(
    "https://geekbyteonline.sharepoint.com/sites/New365",
    "https://geekbyteonline.sharepoint.com/sites/AnotherSite",
    "https://geekbyteonline.sharepoint.com/sites/ThirdSite"
)

# Function to get Azure AD user details (requires Microsoft Graph)
function Get-AzureADUserDetails {
    param($UserPrincipalName)
    
    try {
        # If you have Microsoft Graph module installed
        # $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction SilentlyContinue
        # if ($user) {
        #     return @{
        #         Department = $user.Department
        #         JobTitle = $user.JobTitle
        #         OfficeLocation = $user.OfficeLocation
        #         MobilePhone = $user.MobilePhone
        #     }
        # }
        return @{
            Department = "N/A"
            JobTitle = "N/A"
            OfficeLocation = "N/A"
            MobilePhone = "N/A"
        }
    }
    catch {
        return @{
            Department = "N/A"
            JobTitle = "N/A"
            OfficeLocation = "N/A"
            MobilePhone = "N/A"
        }
    }
}

# Create master CSV
$masterTimestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$masterCsvFile = "SharePoint_Group_Members_Detailed_$masterTimestamp.csv"

# Define headers
$headers = "SiteName,SiteUrl,GroupName,GroupOwner,GroupDescription,UserDisplayName,UserLoginName,UserEmail,UserIsSiteAdmin,UserDepartment,UserJobTitle,UserOfficeLocation,UserMobilePhone,PermissionLevels,GroupMemberCount`n"
[System.IO.File]::WriteAllText($masterCsvFile, $headers)

Write-Host "`n====================================================================" -ForegroundColor Cyan
Write-Host "SHAREPOINT GROUP MEMBERS REPORT (WITH DETAILS)" -ForegroundColor Cyan
Write-Host "Output File: $masterCsvFile" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Cyan

foreach ($siteUrl in $SiteUrls) {
    Write-Host "`n====================================================================" -ForegroundColor Cyan
    Write-Host "CONNECTING TO SITE: $siteUrl" -ForegroundColor Cyan
    Write-Host "====================================================================" -ForegroundColor Cyan
    
    try {
        Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId -ErrorAction Stop
        Write-Host "Successfully connected" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to connect: $_" -ForegroundColor Red
        continue
    }

    try {
        $groups = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/sitegroups" -Method Get -ErrorAction Stop
        $allMembers = @()
        
        foreach ($group in $groups.value) {
            Write-Host "Processing Group: $($group.Title)" -ForegroundColor Yellow
            
            try {
                $membersUrl = "$siteUrl/_api/web/sitegroups($($group.Id))/users"
                $members = Invoke-PnPSPRestMethod -Url $membersUrl -Method Get -ErrorAction Stop
                
                $permissionsUrl = "$siteUrl/_api/web/sitegroups($($group.Id))/roledefinitionbindings"
                $permissions = Invoke-PnPSPRestMethod -Url $permissionsUrl -Method Get -ErrorAction Stop
                
                $permLevels = @()
                if ($permissions.value) {
                    $permLevels = $permissions.value | ForEach-Object { $_.Name }
                }
                $permText = if ($permLevels.Count -gt 0) { $permLevels -join " | " } else { "Inherited/No direct permissions" }
                
                foreach ($member in $members.value) {
                    # Get Azure AD details
                    $userPrincipal = $member.LoginName
                    $adDetails = Get-AzureADUserDetails -UserPrincipalName $userPrincipal
                    
                    $memberObj = [PSCustomObject]@{
                        SiteName = $siteUrl.Split("/")[-1]
                        SiteUrl = $siteUrl
                        GroupName = $group.Title
                        GroupOwner = if ($group.OwnerTitle) { $group.OwnerTitle } else { "Not specified" }
                        GroupDescription = if ($group.Description) { $group.Description } else { "No description" }
                        UserDisplayName = if ($member.Title) { $member.Title } else { $member.LoginName }
                        UserLoginName = $member.LoginName
                        UserEmail = if ($member.Email) { $member.Email } else { "No email" }
                        UserIsSiteAdmin = if ($member.IsSiteAdmin) { "Yes" } else { "No" }
                        UserDepartment = $adDetails.Department
                        UserJobTitle = $adDetails.JobTitle
                        UserOfficeLocation = $adDetails.OfficeLocation
                        UserMobilePhone = $adDetails.MobilePhone
                        PermissionLevels = $permText
                        GroupMemberCount = $members.value.Count
                    }
                    
                    $allMembers += $memberObj
                }
                
                Write-Host "  Added $($members.value.Count) members" -ForegroundColor Green
            }
            catch {
                Write-Host "  Error: $_" -ForegroundColor Red
            }
        }
        
        if ($allMembers.Count -gt 0) {
            $allMembers | Export-Csv -Path $masterCsvFile -Append -NoTypeInformation -Encoding UTF8
            Write-Host "`n✅ Exported $($allMembers.Count) members from $($siteUrl.Split("/")[-1])" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Error: $_" -ForegroundColor Red
    }
    
    Disconnect-PnPOnline
}

Write-Host "`n====================================================================" -ForegroundColor Green
Write-Host "✅ REPORT COMPLETED" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Green
Write-Host "File: $masterCsvFile" -ForegroundColor Green
$fileInfo = Get-Item $masterCsvFile
Write-Host "Size: $([math]::Round($fileInfo.Length/1KB, 2)) KB" -ForegroundColor White
Write-Host "====================================================================" -ForegroundColor Green
