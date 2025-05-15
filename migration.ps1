# Authentication Parameters
$TenantId = "0e439a1f-a497-462b-9e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ThumbPrint = "B799789F78628CAE56B4D0F380FD551EB754E0DB"

# Site URLs to compare
$Site1Url = "https://geekbyteonline.sharepoint.com/sites/Site1"
$Site2Url = "https://geekbyteonline.sharepoint.com/sites/Site2"

# Output file
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outputFile = "SiteComparison_REST_$timestamp.csv"

# Connect to SharePoint
try {
    Connect-PnPOnline -Url $Site1Url -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId
    Write-Host "Connected to SharePoint" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect: $_" -ForegroundColor Red
    exit
}

# Function to get all files from a site using REST
function Get-AllSiteFilesViaRest($siteUrl) {
    # Switch connection to the target site
    Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId
    
    $allFiles = @()
    
    # Get all document libraries (excluding hidden/system lists)
    $lists = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists?`$filter=Hidden eq false and BaseTemplate eq 101" -Method Get
    
    foreach ($list in $lists.value) {
        Write-Host "Processing library: $($list.Title)" -ForegroundColor Cyan
        
        # Get all items in the library
        $nextPageUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items?`$select=FileLeafRef,FileRef,File_x0020_Size,Modified,Created,File,FileSystemObjectType,UniqueId&`$expand=File&`$top=1000"
        
        do {
            try {
                $response = Invoke-PnPSPRestMethod -Url $nextPageUrl -Method Get
                $items = $response.value
                $nextPageUrl = $response.'odata.nextLink'
                
                foreach ($item in $items) {
                    if ($item.FileSystemObjectType -eq 0) { # 0 = File
                        $fileInfo = [PSCustomObject]@{
                            Library       = $list.Title
                            FileName     = $item.FileLeafRef
                            ServerPath    = $item.FileRef
                            Size          = if ($item.File_x0020_Size) { [math]::Round($item.File_x0020_Size / 1KB, 2) } else { 0 }
                            Modified      = $item.Modified
                            Created       = $item.Created
                            UniqueId      = $item.UniqueId
                            SiteUrl       = $siteUrl
                        }
                        $allFiles += $fileInfo
                    }
                }
            }
            catch {
                Write-Host "Error retrieving items from $($list.Title): $_" -ForegroundColor Red
                $nextPageUrl = $null
            }
        } while ($nextPageUrl)
    }
    
    return $allFiles
}

# Get files from both sites
Write-Host "Collecting files from Site 1..." -ForegroundColor Yellow
$site1Files = Get-AllSiteFilesViaRest -siteUrl $Site1Url

Write-Host "Collecting files from Site 2..." -ForegroundColor Yellow
$site2Files = Get-AllSiteFilesViaRest -siteUrl $Site2Url

# Compare the files
$comparisonResults = @()

# Find files in Site1 but not in Site2
foreach ($file in $site1Files) {
    $matchingFile = $site2Files | Where-Object { $_.ServerPath -eq $file.ServerPath }
    
    if (-not $matchingFile) {
        $comparisonResults += [PSCustomObject]@{
            FileName      = $file.FileName
            Library       = $file.Library
            ServerPath    = $file.ServerPath
            Status        = "Only in Site1"
            Site1Size     = "$($file.Size) KB"
            Site2Size     = ""
            Site1Modified = $file.Modified
            Site2Modified = ""
            Site1Url      = $file.SiteUrl
            Site2Url      = ""
        }
    }
}

# Find files in Site2 but not in Site1
foreach ($file in $site2Files) {
    $matchingFile = $site1Files | Where-Object { $_.ServerPath -eq $file.ServerPath }
    
    if (-not $matchingFile) {
        $comparisonResults += [PSCustomObject]@{
            FileName      = $file.FileName
            Library       = $file.Library
            ServerPath    = $file.ServerPath
            Status        = "Only in Site2"
            Site1Size     = ""
            Site2Size     = "$($file.Size) KB"
            Site1Modified = ""
            Site2Modified = $file.Modified
            Site1Url      = ""
            Site2Url      = $file.SiteUrl
        }
    }
}

# Find files in both sites but with differences
foreach ($file in $site1Files) {
    $matchingFile = $site2Files | Where-Object { $_.ServerPath -eq $file.ServerPath }
    
    if ($matchingFile) {
        $differences = @()
        
        if ($file.Size -ne $matchingFile.Size) {
            $differences += "Size (Site1: $($file.Size) KB vs Site2: $($matchingFile.Size) KB)"
        }
        
        if ($file.Modified -ne $matchingFile.Modified) {
            $timeDiff = New-TimeSpan -Start $file.Modified -End $matchingFile.Modified
            $differences += "Modified date differs by $($timeDiff.TotalHours.ToString('0.00')) hours"
        }
        
        if ($differences.Count -gt 0) {
            $comparisonResults += [PSCustomObject]@{
                FileName      = $file.FileName
                Library       = $file.Library
                ServerPath    = $file.ServerPath
                Status        = "Different - " + ($differences -join ", ")
                Site1Size     = "$($file.Size) KB"
                Site2Size     = "$($matchingFile.Size) KB"
                Site1Modified = $file.Modified
                Site2Modified = $matchingFile.Modified
                Site1Url      = $file.SiteUrl
                Site2Url      = $matchingFile.SiteUrl
            }
        }
    }
}

# Export results
if ($comparisonResults.Count -gt 0) {
    $comparisonResults | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
    Write-Host "Comparison completed. Results saved to $outputFile" -ForegroundColor Green
    
    # Display summary
    $onlyInSite1 = ($comparisonResults | Where-Object { $_.Status -eq "Only in Site1" }).Count
    $onlyInSite2 = ($comparisonResults | Where-Object { $_.Status -eq "Only in Site2" }).Count
    $different = ($comparisonResults | Where-Object { $_.Status -like "Different*" }).Count
    
    Write-Host "`nComparison Summary:"
    Write-Host "Files only in $($Site1Url.Split('/')[-1]): $onlyInSite1"
    Write-Host "Files only in $($Site2Url.Split('/')[-1]): $onlyInSite2"
    Write-Host "Files with differences: $different"
}
else {
    Write-Host "No differences found between the sites." -ForegroundColor Yellow
}

# Disconnect
Disconnect-PnPOnline