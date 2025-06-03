<#
.SYNOPSIS
    Recursively uploads files to SharePoint with resume capability and progress tracking.
.DESCRIPTION
    Enhanced version that tracks completed files and can resume interrupted uploads.
    Generates detailed reports and only uploads pending files on subsequent runs.
	
	.\UploadToSharePoint.ps1 -ClientID "your-id" -TenantID "your-tenant" `
    -ClientSecret "your-secret" -SiteId "site-id" -DriveId "drive-id" `
    -SourcePath "C:\YourFiles" -DestinationRootPath "Documents/Uploads"
	
	.\UploadToSharePoint.ps1 ... -DebugMode
	
	.\UploadToSharePoint.ps1 ... -StateFile "CustomState.json"
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$ClientID,
    
    [Parameter(Mandatory=$true)]
    [string]$TenantID,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory=$true)]
    [string]$SiteId,
    
    [Parameter(Mandatory=$true)]
    [string]$DriveId,
    
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,
    
    [Parameter(Mandatory=$true)]
    [string]$DestinationRootPath,
    
    [Parameter(Mandatory=$false)]
    [int]$ChunkSize = 10MB,
    
    [Parameter(Mandatory=$false)]
    [switch]$DebugMode,
    
    [Parameter(Mandatory=$false)]
    [string]$StateFile = "UploadState.json"
)

# Constants
$GraphApiVersion = "v1.0"
$MaxSimpleUploadSize = 4MB
$AllowedRetries = 3

# Initialize tracking objects
$script:SessionState = @{
    TotalFiles = 0
    CompletedFiles = @()
    PendingFiles = @()
    FailedFiles = @()
    StartTime = Get-Date
}

# Load previous state if exists
if (Test-Path $StateFile) {
    try {
        $previousState = Get-Content $StateFile | ConvertFrom-Json -AsHashtable
        $script:SessionState.CompletedFiles = $previousState.CompletedFiles
        Write-Host "Resuming previous upload session..." -ForegroundColor Yellow
        Write-Host "Found $($script:SessionState.CompletedFiles.Count) previously completed files" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to load previous state file: $_"
    }
}

# Enhanced URL encoding function
function Encode-SharePointUrl {
    param (
        [string]$Path
    )
    
    $Path = $Path.Replace('\', '/')
    $segments = $Path -split '/' | Where-Object { $_ -ne '' }
    $encodedSegments = foreach ($segment in $segments) {
        [System.Web.HttpUtility]::UrlEncode($segment).Replace('+', '%20')
    }
    $encodedPath = $encodedSegments -join '/'
    if ($Path.StartsWith('/')) { $encodedPath = "/$encodedPath" }
    return $encodedPath
}

function Write-DebugLog {
    param (
        [string]$Message
    )
    if ($DebugMode) {
        Write-Host "[DEBUG] $(Get-Date -Format 'HH:mm:ss'): $Message" -ForegroundColor Cyan
    }
}

function Save-SessionState {
    param (
        [string]$FilePath
    )
    try {
        $script:SessionState.EndTime = Get-Date
        $script:SessionState.Duration = ($script:SessionState.EndTime - $script:SessionState.StartTime).ToString()
        $script:SessionState | ConvertTo-Json -Depth 5 | Out-File $FilePath
        Write-DebugLog "Session state saved to $FilePath"
    }
    catch {
        Write-Warning "Failed to save session state: $_"
    }
}

function Get-AccessToken {
    try {
        $tokenUrl = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
        $tokenBody = @{
            client_id     = $ClientID
            scope        = "https://graph.microsoft.com/.default"
            client_secret = $ClientSecret
            grant_type    = "client_credentials"
        }

        Write-DebugLog "Requesting access token"
        $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $tokenBody
        return $tokenResponse.access_token
    }
    catch {
        Write-Error "Token request failed: $_"
        exit 1
    }
}

function Ensure-SharePointFolder {
    param (
        [string]$AccessToken,
        [string]$FolderPath
    )

    $retryCount = 0
    while ($retryCount -lt $AllowedRetries) {
        try {
            $encodedPath = Encode-SharePointUrl -Path $FolderPath
            $graphUrl = "https://graph.microsoft.com/$GraphApiVersion/sites/$SiteId/drives/$DriveId/root:/$encodedPath"
            
            Write-DebugLog "Checking folder existence: $graphUrl"
            $headers = @{
                "Authorization" = "Bearer $AccessToken"
                "Content-Type" = "application/json"
            }

            $response = Invoke-RestMethod -Uri $graphUrl -Method Get -Headers $headers -ErrorAction Stop
            return $response
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 404) {
                $parentPath = [System.IO.Path]::GetDirectoryName($FolderPath.TrimEnd('/'))
                $folderName = [System.IO.Path]::GetFileName($FolderPath.TrimEnd('/'))

                if ([string]::IsNullOrEmpty($parentPath)) {
                    $parentPath = " "
                }

                $parentUrl = "https://graph.microsoft.com/$GraphApiVersion/sites/$SiteId/drives/$DriveId/root:/$(Encode-SharePointUrl -Path $parentPath)"
                Write-DebugLog "Getting parent folder: $parentUrl"
                $parentResponse = Invoke-RestMethod -Uri $parentUrl -Method Get -Headers $headers

                $createUrl = "https://graph.microsoft.com/$GraphApiVersion/sites/$SiteId/drives/$DriveId/items/$($parentResponse.id)/children"
                $body = @{
                    "name" = $folderName
                    "folder" = @{}
                    "@microsoft.graph.conflictBehavior" = "rename"
                } | ConvertTo-Json

                Write-DebugLog "Creating folder: $createUrl"
                return Invoke-RestMethod -Uri $createUrl -Method Post -Headers $headers -Body $body -ContentType "application/json"
            }
            else {
                $retryCount++
                if ($retryCount -ge $AllowedRetries) {
                    throw "Folder operation failed after $AllowedRetries attempts: $_"
                }
                Start-Sleep -Seconds (5 * $retryCount)
            }
        }
    }
}

function Upload-File {
    param (
        [string]$AccessToken,
        [string]$FilePath,
        [string]$DestinationPath
    )

    $fileSize = (Get-Item $FilePath).Length
    $fileName = [System.IO.Path]::GetFileName($FilePath)
    $encodedPath = Encode-SharePointUrl -Path $DestinationPath

    try {
        if ($fileSize -le $MaxSimpleUploadSize) {
            $graphUrl = "https://graph.microsoft.com/$GraphApiVersion/sites/$SiteId/drives/$DriveId/root:/${encodedPath}:/content"
            Write-DebugLog "Simple upload URL: $graphUrl"

            $headers = @{
                "Authorization" = "Bearer $AccessToken"
                "Content-Type" = "application/octet-stream"
            }

            $fileContent = [System.IO.File]::ReadAllBytes($FilePath)
            $response = Invoke-RestMethod -Uri $graphUrl -Method Put -Headers $headers -Body $fileContent -ErrorAction Stop
            
            # Mark as completed
            $script:SessionState.CompletedFiles += @{
                Source = $FilePath
                Destination = $DestinationPath
                Size = $fileSize
                Timestamp = Get-Date
            }
            Save-SessionState -FilePath $StateFile
            
            Write-Host "✓ Uploaded: $DestinationPath" -ForegroundColor Green
            return $response
        }
        else {
            $sessionUrl = "https://graph.microsoft.com/$GraphApiVersion/sites/$SiteId/drives/$DriveId/root:/${encodedPath}:/createUploadSession"
            Write-DebugLog "Creating upload session: $sessionUrl"

            $headers = @{
                "Authorization" = "Bearer $AccessToken"
                "Content-Type" = "application/json"
            }

            $session = Invoke-RestMethod -Uri $sessionUrl -Method Post -Headers $headers -ErrorAction Stop
            $uploadUrl = $session.uploadUrl

            $fileStream = [System.IO.File]::OpenRead($FilePath)
            $reader = New-Object System.IO.BinaryReader($fileStream)
            $buffer = New-Object byte[] $ChunkSize

            $bytesUploaded = 0
            while ($bytesUploaded -lt $fileSize) {
                $remaining = $fileSize - $bytesUploaded
                $currentChunkSize = [Math]::Min($ChunkSize, $remaining)
                $bytesRead = $reader.Read($buffer, 0, $currentChunkSize)

                $chunkContent = New-Object byte[] $currentChunkSize
                [Array]::Copy($buffer, $chunkContent, $currentChunkSize)

                $retryCount = 0
                $success = $false
                while (-not $success -and $retryCount -lt $AllowedRetries) {
                    try {
                        $headers = @{
                            "Authorization" = "Bearer $AccessToken"
                            "Content-Length" = $currentChunkSize
                            "Content-Range" = "bytes $bytesUploaded-$($bytesUploaded + $currentChunkSize - 1)/$fileSize"
                        }

                        Write-DebugLog "Uploading chunk $($bytesUploaded)-$($bytesUploaded + $currentChunkSize - 1)"
                        $null = Invoke-RestMethod -Uri $uploadUrl -Method Put -Headers $headers -Body $chunkContent -ErrorAction Stop
                        
                        $bytesUploaded += $currentChunkSize
                        $success = $true
                        Write-Progress -Activity "Uploading $fileName" -Status "$([math]::Round($bytesUploaded/1MB,2)) MB of $([math]::Round($fileSize/1MB,2)) MB" -PercentComplete (($bytesUploaded / $fileSize) * 100)
                    }
                    catch {
                        $retryCount++
                        if ($retryCount -ge $AllowedRetries) {
                            throw "Chunk upload failed after $AllowedRetries attempts: $_"
                        }
                        Start-Sleep -Seconds (2 * $retryCount)
                    }
                }
            }

            $reader.Close()
            $fileStream.Close()
            
            # Mark as completed
            $script:SessionState.CompletedFiles += @{
                Source = $FilePath
                Destination = $DestinationPath
                Size = $fileSize
                Timestamp = Get-Date
            }
            Save-SessionState -FilePath $StateFile
            
            Write-Host "✓ Uploaded (chunked): $DestinationPath" -ForegroundColor Green
            return $session
        }
    }
    catch {
        # Mark as failed
        $script:SessionState.FailedFiles += @{
            Source = $FilePath
            Destination = $DestinationPath
            Error = $_.Exception.Message
            Timestamp = Get-Date
        }
        Save-SessionState -FilePath $StateFile
        
        Write-Host "✗ Failed to upload $DestinationPath" -ForegroundColor Red
        Write-Host "Error details: $_" -ForegroundColor Yellow
        Write-DebugLog "Failed URL: $graphUrl"
        return $null
    }
    finally {
        if ($reader -ne $null) { $reader.Close() }
        if ($fileStream -ne $null) { $fileStream.Close() }
    }
}

function Get-Report {
    $report = @"
=============================================
         SharePoint Upload Report
=============================================
Start Time:    $($script:SessionState.StartTime)
End Time:      $(if ($script:SessionState.EndTime) { $script:SessionState.EndTime } else { "In Progress" })
Duration:      $(if ($script:SessionState.Duration) { $script:SessionState.Duration } else { "$((Get-Date) - $script:SessionState.StartTime)" })

Total Files:   $($script:SessionState.TotalFiles)
Completed:     $($script:SessionState.CompletedFiles.Count)
Pending:       $($script:SessionState.PendingFiles.Count)
Failed:        $($script:SessionState.FailedFiles.Count)

Source Path:   $SourcePath
Destination:   $DestinationRootPath
"@

    if ($script:SessionState.FailedFiles.Count -gt 0) {
        $report += "`n`nFailed Files:`n"
        $report += ($script:SessionState.FailedFiles | ForEach-Object { 
            " - $($_.Source) -> $($_.Destination)`n   Error: $($_.Error)`n" 
        }) -join "`n"
    }

    return $report
}

# Main execution
try {
    $accessToken = Get-AccessToken
    $allFiles = Get-ChildItem -Path $SourcePath -File -Recurse
    $script:SessionState.TotalFiles = $allFiles.Count

    # Filter out already completed files
    $pendingFiles = $allFiles | Where-Object {
        $filePath = $_.FullName
        -not ($script:SessionState.CompletedFiles | Where-Object { $_.Source -eq $filePath })
    }
    $script:SessionState.PendingFiles = $pendingFiles.Count

    Write-Host "Starting upload session..." -ForegroundColor Cyan
    Write-Host "Total files: $($allFiles.Count)" -ForegroundColor White
    Write-Host "Already completed: $($script:SessionState.CompletedFiles.Count)" -ForegroundColor Green
    Write-Host "Pending upload: $($pendingFiles.Count)" -ForegroundColor Yellow

    foreach ($file in $pendingFiles) {
        $relativePath = $file.FullName.Substring($SourcePath.Length).TrimStart('\')
        $destinationPath = "$DestinationRootPath/$relativePath".Replace('\', '/')
        $destinationFolder = [System.IO.Path]::GetDirectoryName($destinationPath).Replace('\', '/')

        try {
            # Ensure folder structure exists
            if (-not [string]::IsNullOrEmpty($destinationFolder)) {
                Write-DebugLog "Ensuring folder: $destinationFolder"
                $null = Ensure-SharePointFolder -AccessToken $accessToken -FolderPath $destinationFolder
            }

            # Upload the file
            Write-DebugLog "Processing file: $($file.FullName) -> $destinationPath"
            $result = Upload-File -AccessToken $accessToken -FilePath $file.FullName -DestinationPath $destinationPath

            if (-not $result) {
                Write-DebugLog "Upload failed for $destinationPath"
            }
        }
        catch {
            Write-Host "✗ Error processing $($file.FullName): $_" -ForegroundColor Red
            continue
        }
    }

    # Generate and display final report
    $report = Get-Report
    Write-Host $report
    
    # Save report to file
    $report | Out-File "UploadReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    Write-Host "`nReport saved to UploadReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt" -ForegroundColor Cyan
    
    if ($script:SessionState.FailedFiles.Count -gt 0) {
        exit 1
    }
}
catch {
    Write-Error "Fatal error: $_"
    Save-SessionState -FilePath $StateFile
    exit 1
}
finally {
    Save-SessionState -FilePath $StateFile
}
