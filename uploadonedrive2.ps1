<#
.SYNOPSIS
    Recursively uploads files to SharePoint with detailed CSV reporting and resume capability.
.DESCRIPTION
    Enhanced version with CSV tracking, source file metadata, and dynamic status updates.
    Supports both direct access token and client credentials authentication.
#>

param (
    # Authentication options
    [Parameter(ParameterSetName='TokenAuth', Mandatory=$true)]
    [string]$AccessToken,
    
    [Parameter(ParameterSetName='ClientAuth', Mandatory=$true)]
    [string]$ClientID,
    
    [Parameter(ParameterSetName='ClientAuth', Mandatory=$true)]
    [string]$TenantID,
    
    [Parameter(ParameterSetName='ClientAuth', Mandatory=$true)]
    [string]$ClientSecret,
    
    # SharePoint configuration
    [Parameter(Mandatory=$true)]
    [string]$SiteId,
    
    [Parameter(Mandatory=$true)]
    [string]$DriveId,
    
    # File operations
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,
    
    [Parameter(Mandatory=$true)]
    [string]$DestinationRootPath,
    
    [Parameter(Mandatory=$false)]
    [int]$ChunkSize = 10MB,
    
    # Reporting options
    [Parameter(Mandatory=$false)]
    [string]$ReportFile = "UploadReport.csv",
    
    [Parameter(Mandatory=$false)]
    [string]$StateFile = "UploadState.json",
    
    # Debugging
    [Parameter(Mandatory=$false)]
    [switch]$DebugMode
)

# Constants
$GraphApiVersion = "v1.0"
$MaxSimpleUploadSize = 4MB
$AllowedRetries = 3

# Initialize tracking objects
$script:SessionState = @{
    Version = "3.0"
    TotalFiles = 0
    CompletedFiles = @()
    FailedFiles = @()
    StartTime = Get-Date
    SourcePath = $SourcePath
    DestinationRoot = $DestinationRootPath
    ReportFile = $ReportFile
}

# Initialize CSV report structure
$script:CsvReport = @()
$script:ReportLock = [System.Threading.ReaderWriterLockSlim]::new()

# Load previous state if exists
function Initialize-Report {
    param (
        [string]$ReportPath,
        [string]$SourcePath
    )

    # Create new report if doesn't exist
    if (-not (Test-Path $ReportPath)) {
        $reportHeader = "SourcePath,FileName,FileSize,LastModified,UploadStatus,DestinationPath,Timestamp,Error"
        $reportHeader | Out-File $ReportPath -Force
        Write-DebugLog "Created new report file at $ReportPath"
        return @()
    }

    # Load existing report - Fixed handling of CSV import
    try {
        $existingReport = Import-Csv $ReportPath | ForEach-Object {
            [PSCustomObject]@{
                SourcePath = $_.SourcePath
                FileName = $_.FileName
                FileSize = [long]$_.FileSize
                LastModified = [datetime]$_.LastModified
                UploadStatus = $_.UploadStatus
                DestinationPath = $_.DestinationPath
                Timestamp = $_.Timestamp
                Error = $_.Error
            }
        }
        Write-DebugLog "Loaded existing report with $($existingReport.Count) entries"
        return @($existingReport)  # Ensure we return an array
    }
    catch {
        Write-Warning "Failed to load existing report: $_"
        return @()
    }
}

# Load previous state if exists
if (Test-Path $StateFile) {
    try {
        $previousState = Get-Content $StateFile | ConvertFrom-Json -AsHashtable
        
        # Only use previous state if it matches current source and destination
        if ($previousState.SourcePath -eq $SourcePath -and $previousState.DestinationRoot -eq $DestinationRootPath) {
            $script:SessionState.CompletedFiles = $previousState.CompletedFiles
            $script:SessionState.FailedFiles = $previousState.FailedFiles
            Write-Host "Resuming previous upload session..." -ForegroundColor Yellow
            Write-Host "Found $($script:SessionState.CompletedFiles.Count) previously completed files" -ForegroundColor Green
            Write-Host "Found $($script:SessionState.FailedFiles.Count) previously failed files" -ForegroundColor Red
        }
        else {
            Write-Host "Previous state doesn't match current parameters - starting fresh upload" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Warning "Failed to load previous state file: $_"
    }
}

# Initialize CSV report
$script:CsvReport = Initialize-Report -ReportPath $ReportFile -SourcePath $SourcePath

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

function Update-CsvReport {
    param (
        [string]$SourcePath,
        [string]$FileName,
        [long]$FileSize,
        [datetime]$LastModified,
        [string]$UploadStatus,
        [string]$DestinationPath,
        [string]$ErrorDetails = ""
    )

    try {
        $script:ReportLock.EnterWriteLock()
        
        $timestamp = (Get-Date).ToString("o")
        $newEntry = [PSCustomObject]@{
            SourcePath = $SourcePath
            FileName = $FileName
            FileSize = $FileSize
            LastModified = $LastModified.ToString("o")
            UploadStatus = $UploadStatus
            DestinationPath = $DestinationPath
            Timestamp = $timestamp
            Error = $ErrorDetails
        }

        # Check if file already exists in report
        $existingEntry = $script:CsvReport | Where-Object { $_.SourcePath -eq $SourcePath } | Select-Object -First 1
        
        if ($existingEntry) {
            # Update existing entry
            $existingEntry.FileName = $FileName
            $existingEntry.FileSize = $FileSize
            $existingEntry.LastModified = $LastModified.ToString("o")
            $existingEntry.UploadStatus = $UploadStatus
            $existingEntry.DestinationPath = $DestinationPath
            $existingEntry.Timestamp = $timestamp
            $existingEntry.Error = $ErrorDetails
        }
        else {
            # Add new entry - Fixed array addition
            $script:CsvReport = @($script:CsvReport) + $newEntry
        }

        # Update CSV file - Fixed export to handle all cases
        $script:CsvReport | Export-Csv $ReportFile -NoTypeInformation -Force
        
        Write-DebugLog "Updated report for $FileName - Status: $UploadStatus"
    }
    catch {
        Write-Warning "Failed to update CSV report: $_"
        Write-Warning "Error details: $($_.Exception.ToString())"
    }
    finally {
        if ($script:ReportLock.IsWriteLockHeld) {
            $script:ReportLock.ExitWriteLock()
        }
    }
}

function Save-SessionState {
    param (
        [string]$FilePath
    )
    try {
        $script:SessionState.EndTime = Get-Date
        $script:SessionState.Duration = ($script:SessionState.EndTime - $script:SessionState.StartTime).ToString()
        $script:SessionState | ConvertTo-Json -Depth 5 | Out-File $FilePath -Force
        Write-DebugLog "Session state saved to $FilePath"
    }
    catch {
        Write-Warning "Failed to save session state: $_"
    }
}

function Get-AccessToken {
    # If AccessToken was provided directly, use that
    if ($PSCmdlet.ParameterSetName -eq 'TokenAuth') {
        Write-DebugLog "Using provided access token"
        return $AccessToken
    }
    
    # Otherwise get new token using client credentials
    try {
        $tokenUrl = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
        $tokenBody = @{
            client_id     = $ClientID
            scope        = "https://graph.microsoft.com/.default"
            client_secret = $ClientSecret
            grant_type    = "client_credentials"
        }

        Write-DebugLog "Requesting new access token"
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
        [System.IO.FileInfo]$File,
        [string]$DestinationPath
    )

    $fileSize = $File.Length
    $fileName = $File.Name
    $filePath = $File.FullName
    $lastModified = $File.LastWriteTime
    $encodedPath = Encode-SharePointUrl -Path $DestinationPath

    # Update report with "In Progress" status
    Update-CsvReport -SourcePath $filePath -FileName $fileName -FileSize $fileSize `
        -LastModified $lastModified -UploadStatus "In Progress" -DestinationPath $DestinationPath

    try {
        if ($fileSize -le $MaxSimpleUploadSize) {
            $graphUrl = "https://graph.microsoft.com/$GraphApiVersion/sites/$SiteId/drives/$DriveId/root:/${encodedPath}:/content"
            Write-DebugLog "Simple upload URL: $graphUrl"

            $headers = @{
                "Authorization" = "Bearer $AccessToken"
                "Content-Type" = "application/octet-stream"
            }

            $fileContent = [System.IO.File]::ReadAllBytes($filePath)
            $response = Invoke-RestMethod -Uri $graphUrl -Method Put -Headers $headers -Body $fileContent -ErrorAction Stop
            
            # Mark as completed in session state
            $completedFile = @{
                Source = $filePath
                Destination = $DestinationPath
                Size = $fileSize
                LastModified = $lastModified.ToString("o")
                Timestamp = (Get-Date).ToString("o")
            }
            $script:SessionState.CompletedFiles += $completedFile
            
            # Remove from failed files if it was there before
            $script:SessionState.FailedFiles = @($script:SessionState.FailedFiles | Where-Object { $_.Source -ne $filePath })
            
            Save-SessionState -FilePath $StateFile
            
            # Update report with success
            Update-CsvReport -SourcePath $filePath -FileName $fileName -FileSize $fileSize `
                -LastModified $lastModified -UploadStatus "Completed" -DestinationPath $DestinationPath
            
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

            $fileStream = [System.IO.File]::OpenRead($filePath)
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
                        
                        # Update progress percentage in CSV
                        $progress = [math]::Round(($bytesUploaded / $fileSize) * 100)
                        Update-CsvReport -SourcePath $filePath -FileName $fileName -FileSize $fileSize `
                            -LastModified $lastModified -UploadStatus "Uploading ($progress%)" -DestinationPath $DestinationPath
                            
                        Write-Progress -Activity "Uploading $fileName" -Status "$([math]::Round($bytesUploaded/1MB,2)) MB of $([math]::Round($fileSize/1MB,2)) MB" -PercentComplete $progress
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
            
            # Mark as completed in session state
            $completedFile = @{
                Source = $filePath
                Destination = $DestinationPath
                Size = $fileSize
                LastModified = $lastModified.ToString("o")
                Timestamp = (Get-Date).ToString("o")
            }
            $script:SessionState.CompletedFiles += $completedFile
            
            # Remove from failed files if it was there before
            $script:SessionState.FailedFiles = @($script:SessionState.FailedFiles | Where-Object { $_.Source -ne $filePath })
            
            Save-SessionState -FilePath $StateFile
            
            # Update report with success
            Update-CsvReport -SourcePath $filePath -FileName $fileName -FileSize $fileSize `
                -LastModified $lastModified -UploadStatus "Completed" -DestinationPath $DestinationPath
            
            Write-Host "✓ Uploaded (chunked): $DestinationPath" -ForegroundColor Green
            return $session
        }
    }
    catch {
        # Mark as failed (only if not already in completed list)
        if (-not ($script:SessionState.CompletedFiles | Where-Object { $_.Source -eq $filePath })) {
            $failedFile = @{
                Source = $filePath
                Destination = $DestinationPath
                Error = $_.Exception.Message
                LastModified = $lastModified.ToString("o")
                Timestamp = (Get-Date).ToString("o")
            }
            $script:SessionState.FailedFiles = @($script:SessionState.FailedFiles | Where-Object { $_.Source -ne $filePath }) + $failedFile
            Save-SessionState -FilePath $StateFile
            
            # Update report with failure
            Update-CsvReport -SourcePath $filePath -FileName $fileName -FileSize $fileSize `
                -LastModified $lastModified -UploadStatus "Failed" -DestinationPath $DestinationPath -ErrorDetails $_.Exception.Message
        }
        
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
    $pendingCount = $script:SessionState.TotalFiles - $script:SessionState.CompletedFiles.Count - $script:SessionState.FailedFiles.Count
    $pendingCount = [Math]::Max(0, $pendingCount)  # Ensure not negative
    
    $report = @"
=============================================
         SharePoint Upload Report
=============================================
Start Time:    $($script:SessionState.StartTime)
End Time:      $(if ($script:SessionState.EndTime) { $script:SessionState.EndTime } else { "In Progress" })
Duration:      $(if ($script:SessionState.Duration) { $script:SessionState.Duration } else { "$((Get-Date) - $script:SessionState.StartTime)" })

Source Path:   $SourcePath
Destination:   $DestinationRootPath

Total Files:   $($script:SessionState.TotalFiles)
Completed:     $($script:SessionState.CompletedFiles.Count) ($([math]::Round(($script:SessionState.CompletedFiles.Count/$script:SessionState.TotalFiles*100), 1))%)
Pending:       $pendingCount ($([math]::Round(($pendingCount/$script:SessionState.TotalFiles*100), 1))%)
Failed:        $($script:SessionState.FailedFiles.Count) ($([math]::Round(($script:SessionState.FailedFiles.Count/$script:SessionState.TotalFiles*100), 1))%)

Report File:   $ReportFile
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
    # Get access token (either provided directly or via client credentials)
    $accessToken = Get-AccessToken
    
    # Get all files and filter out completed ones
    $allFiles = @(Get-ChildItem -Path $SourcePath -File -Recurse)
    $script:SessionState.TotalFiles = $allFiles.Count
    
    # Calculate pending files (total - completed - failed)
    $pendingFiles = @($allFiles | Where-Object {
        $filePath = $_.FullName
        -not ($script:SessionState.CompletedFiles | Where-Object { $_.Source -eq $filePath }) -and
        -not ($script:SessionState.FailedFiles | Where-Object { $_.Source -eq $filePath })
    })
    
    Write-Host "Starting upload session..." -ForegroundColor Cyan
    Write-Host "Total files: $($allFiles.Count)" -ForegroundColor White
    Write-Host "Already completed: $($script:SessionState.CompletedFiles.Count)" -ForegroundColor Green
    Write-Host "Pending upload: $($pendingFiles.Count)" -ForegroundColor Yellow
    Write-Host "Previously failed: $($script:SessionState.FailedFiles.Count)" -ForegroundColor Red
    Write-Host "Report file: $ReportFile" -ForegroundColor Cyan

    # Initialize CSV report with all files if empty
    if ($script:CsvReport.Count -eq 0) {
        Write-DebugLog "Initializing CSV report with all files"
        foreach ($file in $allFiles) {
            $relativePath = $file.FullName.Substring($SourcePath.Length).TrimStart('\')
            $destinationPath = "$DestinationRootPath/$relativePath".Replace('\', '/')
            
            # Check if already completed in session state
            $status = if ($script:SessionState.CompletedFiles | Where-Object { $_.Source -eq $file.FullName }) {
                "Completed"
            } elseif ($script:SessionState.FailedFiles | Where-Object { $_.Source -eq $file.FullName }) {
                "Failed"
            } else {
                "Pending"
            }
            
            Update-CsvReport -SourcePath $file.FullName -FileName $file.Name -FileSize $file.Length `
                -LastModified $file.LastWriteTime -UploadStatus $status -DestinationPath $destinationPath
        }
    }

    # Process pending files
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
            $result = Upload-File -AccessToken $accessToken -File $file -DestinationPath $destinationPath

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
