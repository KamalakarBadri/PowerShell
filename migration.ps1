param (
    [string]$SourceSiteUrl = "<SOURCE_SITE_URL>",
    [string]$SourceLibrary = "Documents",
    [string]$DestSiteUrl = "<DEST_SITE_URL>",
    [string]$DestLibrary = "Documents",
    [string]$TenantId = "<TENANT_ID>",
    [string]$ClientId = "<CLIENT_ID>",
    [string]$ClientSecret = "<CLIENT_SECRET>",
    [string]$ReportPath = "SharePointMigrationComparisonReport.csv"
)

function Get-AccessToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Resource
    )
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        resource      = $Resource
    }
    $response = Invoke-RestMethod -Method Post -Uri "https://accounts.accesscontrol.windows.net/$TenantId/tokens/OAuth/2" -Body $body
    return $response.access_token
}

function Get-SPOSiteFiles {
    param (
        [string]$SiteUrl,
        [string]$Library,
        [string]$AccessToken
    )
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Accept"        = "application/json;odata=verbose"
    }
    $files = @()
    $skip = 0
    $pageSize = 5000
    do {
        $url = "$SiteUrl/_api/web/lists/GetByTitle('$Library')/items?`$select=FileRef,FileLeafRef,FSObjType,File_x0020_Size&`$top=$pageSize&`$skip=$skip"
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        foreach ($item in $resp.d.results) {
            if ($item.FSObjType -eq 0) { # 0 = File
                $files += [PSCustomObject]@{
                    Path = $item.FileRef.ToLower()
                    Name = $item.FileLeafRef
                    Size = [int64]($item.File_x0020_Size)
                }
            }
        }
        $skip += $pageSize
    } while ($resp.d.results.Count -eq $pageSize)
    return $files
}

# Get access tokens for both sites (assumes same tenant/app)
$resource = "00000003-0000-0ff1-ce00-000000000000/$($SourceSiteUrl -replace '^https://','')@${TenantId}"
$sourceToken = Get-AccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -Resource $resource

$resourceDest = "00000003-0000-0ff1-ce00-000000000000/$($DestSiteUrl -replace '^https://','')@${TenantId}"
$destToken = Get-AccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -Resource $resourceDest

Write-Host "Retrieving files from source site..." -ForegroundColor Cyan
$sourceFiles = Get-SPOSiteFiles -SiteUrl $SourceSiteUrl -Library $SourceLibrary -AccessToken $sourceToken

Write-Host "Retrieving files from destination site..." -ForegroundColor Cyan
$destFiles = Get-SPOSiteFiles -SiteUrl $DestSiteUrl -Library $DestLibrary -AccessToken $destToken

# Convert to hashtables for fast lookup
$sourceHash = @{}
foreach ($f in $sourceFiles) { $sourceHash[$f.Path] = $f }

$destHash = @{}
foreach ($f in $destFiles) { $destHash[$f.Path] = $f }

$allPaths = ($sourceHash.Keys + $destHash.Keys) | Sort-Object -Unique

$report = @()

foreach ($path in $allPaths) {
    $inSource = $sourceHash.ContainsKey($path)
    $inDest = $destHash.ContainsKey($path)

    if ($inSource -and $inDest) {
        if ($sourceHash[$path].Size -eq $destHash[$path].Size) {
            $status = "Matched"
        } else {
            $status = "SizeMismatch"
        }
        $report += [PSCustomObject]@{
            Path = $path
            SourceSize = $sourceHash[$path].Size
            DestSize = $destHash[$path].Size
            Status = $status
        }
    } elseif ($inSource -and -not $inDest) {
        $report += [PSCustomObject]@{
            Path = $path
            SourceSize = $sourceHash[$path].Size
            DestSize = ""
            Status = "MissingInDestination"
        }
    } elseif (-not $inSource -and $inDest) {
        $report += [PSCustomObject]@{
            Path = $path
            SourceSize = ""
            DestSize = $destHash[$path].Size
            Status = "ExtraInDestination"
        }
    }
}

$report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
Write-Host "Comparison report saved to $ReportPath" -ForegroundColor Green