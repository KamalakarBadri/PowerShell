# PowerShell script to checkout, update fields, and checkin SharePoint files
# Prerequisites: Install-Module -Name PnP.PowerShell -Force

# Configuration - Update these based on your SharePoint environment
$SiteUrl = "https://yourcompany.sharepoint.com/sites/yoursite"
$DocumentLibrary = "Documents" # Your document library name
$FolderPath = "Subfolder1/Subfolder2" # Folder path within the document library (leave empty "" for root)

# Logging configuration
$LogPath = "C:\Logs\SharePoint_FileUpdate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$EnableLogging = $true

# Option 1: Specify files manually (comment out if using Option 2)
$FilesToUpdate = @(
    "Document1.docx",
    "Document2.pdf", 
    "Document3.xlsx",
    "Report.pptx"
    # Add more files as needed
)

# Option 2: Auto-discover files with "In Progress" status (set to $true to enable)
$AutoDiscoverFiles = $true

# Fields to update - These are the three required columns
$FieldsToUpdate = @{
    "CurrentApprovalStatus" = "Approved"           # Current Approval Status
    "FinalApprovalDate" = (Get-Date).ToString("yyyy-MM-dd")  # Final Approval Date
    "DocumentStatus" = "Approved"                  # Document Status
}

# Logging function
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Color = "White"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to console with color
    Write-Host $Message -ForegroundColor $Color
    
    # Write to log file if logging is enabled
    if ($EnableLogging) {
        try {
            # Ensure log directory exists
            $logDir = Split-Path $LogPath -Parent
            if (!(Test-Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }
            
            # Append to log file
            Add-Content -Path $LogPath -Value $logEntry -Encoding UTF8
        }
        catch {
            Write-Host "Warning: Could not write to log file: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# Function to initialize log file
function Initialize-LogFile {
    if ($EnableLogging) {
        try {
            $logDir = Split-Path $LogPath -Parent
            if (!(Test-Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }
            
            $separator = "=" * 80
            $header = @"
$separator
SharePoint File Update Script Log
Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Site URL: $SiteUrl
Library: $DocumentLibrary
Folder: $(if([string]::IsNullOrEmpty($FolderPath)) { "Root" } else { $FolderPath })
Auto-discover: $AutoDiscoverFiles
$separator
"@
            Set-Content -Path $LogPath -Value $header -Encoding UTF8
            Write-Host "Log file initialized: $LogPath" -ForegroundColor Green
        }
        catch {
            Write-Host "Warning: Could not initialize log file: $($_.Exception.Message)" -ForegroundColor Yellow
            $script:EnableLogging = $false
        }
    }
}
function Get-FilesInProgress {
    param(
        [string]$Library,
        [string]$FolderPath
    )
    
    try {
        Write-Host "Searching for files with 'In Progress' status..." -ForegroundColor Cyan
        
        # Build the folder query
        if ([string]::IsNullOrEmpty($FolderPath)) {
            $folderUrl = $Library
        } else {
            $folderUrl = "$Library/$FolderPath"
        }
        
        # Get all items in the folder/library with "In Progress" status
        $camlQuery = @"
<View Scope='RecursiveAll'>
    <Query>
        <Where>
            <Eq>
                <FieldRef Name='DocumentStatus'/>
                <Value Type='Text'>In Progress</Value>
            </Eq>
        </Where>
    </Query>
</View>
"@
        
        # Get items using CAML query
        $items = Get-PnPListItem -List $Library -Query $camlQuery
        
        # Filter by folder if specified
        if (-not [string]::IsNullOrEmpty($FolderPath)) {
            $items = $items | Where-Object { $_.FieldValues.FileDirRef -like "*$FolderPath*" }
        }
        
        $filesInProgress = @()
        foreach ($item in $items) {
            if ($item.FileSystemObjectType -eq "File") {
                $filesInProgress += [PSCustomObject]@{
                    Name = $item.FieldValues.FileLeafRef
                    Path = $item.FieldValues.FileRef
                    CurrentStatus = $item.FieldValues.DocumentStatus
                    ApprovalStatus = $item.FieldValues.CurrentApprovalStatus
                    Id = $item.Id
                }
            }
        }
        
        return $filesInProgress
    }
    catch {
        Write-Host "Error searching for files: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Function to display files for confirmation
function Show-FilesForConfirmation {
    param([array]$Files)
    
    if ($Files.Count -eq 0) {
        Write-Log "No files found with 'In Progress' status." -Level "WARN" -Color "Yellow"
        return $false
    }
    
    Write-Log "" -Level "INFO" -Color "White"
    Write-Log "=== FILES WITH 'IN PROGRESS' STATUS ===" -Level "INFO" -Color "Cyan"
    Write-Log "Found $($Files.Count) file(s):" -Level "INFO" -Color "White"
    
    for ($i = 0; $i -lt $Files.Count; $i++) {
        $file = $Files[$i]
        Write-Log "  $($i + 1). $($file.Name)" -Level "INFO" -Color "White"
        Write-Log "      Current Status: $($file.CurrentStatus)" -Level "INFO" -Color "Gray"
        Write-Log "      Approval Status: $($file.ApprovalStatus)" -Level "INFO" -Color "Gray"
    }
    
    Write-Log "" -Level "INFO" -Color "White"
    Write-Log "These files will be updated to:" -Level "INFO" -Color "Yellow"
    Write-Log "  • Current Approval Status: Approved" -Level "INFO" -Color "Green"
    Write-Log "  • Final Approval Date: $(Get-Date -Format 'yyyy-MM-dd')" -Level "INFO" -Color "Green"
    Write-Log "  • Document Status: Approved" -Level "INFO" -Color "Green"
    Write-Log "  • Version: Major version with approval comment" -Level "INFO" -Color "Green"
    
    return $true
}
function Update-SharePointFiles {
    param(
        [string]$SiteUrl,
        [string]$Library,
        [array]$Files,
        [hashtable]$Fields
    )
    
    try {
        # Connect to SharePoint Online
        Write-Log "Connecting to SharePoint site: $SiteUrl" -Level "INFO" -Color "Cyan"
        Connect-PnPOnline -Url $SiteUrl -Interactive
        
        Write-Log "Connected successfully!" -Level "INFO" -Color "Green"
        Write-Log "Processing $($Files.Count) files..." -Level "INFO" -Color "Cyan"
        
        $successCount = 0
        $errorCount = 0
        
        foreach ($fileItem in $Files) {
            # Handle both file objects and file names
            if ($UseFileObjects) {
                $fileName = $fileItem.Name
                $fileUrl = $fileItem.Path
                $fileId = $fileItem.Id
                Write-Log "" -Level "INFO" -Color "White"
                Write-Log "--- Processing: $fileName (ID: $fileId) ---" -Level "INFO" -Color "Yellow"
            } else {
                $fileName = $fileItem
                # Build the file URL with folder path
                if ([string]::IsNullOrEmpty($FolderPath)) {
                    $fileUrl = "$Library/$fileName"
                } else {
                    $fileUrl = "$Library/$FolderPath/$fileName"
                }
                Write-Log "" -Level "INFO" -Color "White"
                Write-Log "--- Processing: $fileName ---" -Level "INFO" -Color "Yellow"
            }
            
            try {
                # Get file (skip if we already have the file object with ID)
                if ($UseFileObjects) {
                    $file = Get-PnPListItem -List $Library -Id $fileId -ErrorAction Stop
                    Write-Log "✓ File retrieved using List Item ID: $fileId" -Level "SUCCESS" -Color "Green"
                } else {
                    # Build the file URL with folder path
                    if ([string]::IsNullOrEmpty($FolderPath)) {
                        $fileUrl = "$Library/$fileName"
                    } else {
                        $fileUrl = "$Library/$FolderPath/$fileName"
                    }
                    
                    # Try multiple methods to get the file from document library
                    $file = $null
                    
                    # Method 1: Direct file URL approach
                    try {
                        Write-Log "Method 1: Getting file using direct URL: $fileUrl" -Level "DEBUG" -Color "Gray"
                        $file = Get-PnPFile -Url $fileUrl -AsListItem -ErrorAction Stop
                        Write-Log "✓ File found using direct URL method" -Level "SUCCESS" -Color "Green"
                    }
                    catch {
                        Write-Log "Method 1 failed: $($_.Exception.Message)" -Level "WARN" -Color "Yellow"
                        
                        # Method 2: Search in document library folder
                        try {
                            Write-Log "Method 2: Searching in document library folder..." -Level "DEBUG" -Color "Gray"
                            if ([string]::IsNullOrEmpty($FolderPath)) {
                                $folderItems = Get-PnPFolderItem -FolderSiteRelativeUrl $Library -ItemType File
                            } else {
                                $folderItems = Get-PnPFolderItem -FolderSiteRelativeUrl "$Library/$FolderPath" -ItemType File
                            }
                            
                            $targetFile = $folderItems | Where-Object { $_.Name -eq $fileName }
                            if ($targetFile) {
                                $file = Get-PnPFile -Url $targetFile.ServerRelativeUrl -AsListItem -ErrorAction Stop
                                Write-Log "✓ File found using folder search method" -Level "SUCCESS" -Color "Green"
                            } else {
                                Write-Log "File '$fileName' not found in folder" -Level "ERROR" -Color "Red"
                                Write-Log "Available files in folder:" -Level "INFO" -Color "Yellow"
                                foreach ($availableFile in $folderItems | Select-Object -First 10) {
                                    Write-Log "  - $($availableFile.Name)" -Level "INFO" -Color "Gray"
                                }
                                if ($folderItems.Count -gt 10) {
                                    Write-Log "  ... and $($folderItems.Count - 10) more files" -Level "INFO" -Color "Gray"
                                }
                            }
                        }
                        catch {
                            Write-Log "Method 2 also failed: $($_.Exception.Message)" -Level "ERROR" -Color "Red"
                            
                            # Method 3: Try with server relative URL format
                            try {
                                Write-Log "Method 3: Trying server relative URL format..." -Level "DEBUG" -Color "Gray"
                                $siteUrl = Get-PnPWeb | Select-Object -ExpandProperty ServerRelativeUrl
                                $serverRelativeUrl = "$siteUrl/$fileUrl"
                                Write-Log "Attempting server relative URL: $serverRelativeUrl" -Level "DEBUG" -Color "Gray"
                                $file = Get-PnPFile -Url $serverRelativeUrl -AsListItem -ErrorAction Stop
                                Write-Log "✓ File found using server relative URL method" -Level "SUCCESS" -Color "Green"
                            }
                            catch {
                                Write-Log "Method 3 also failed: $($_.Exception.Message)" -Level "ERROR" -Color "Red"
                            }
                        }
                    }
                    
                    if (-not $file) {
                        throw "Could not locate file '$fileName' in document library. Please verify the file name and folder path are correct."
                    }
                }
                
                if ($file) {
                    Write-Log "✓ File found" -Level "INFO" -Color "Green"
                    
                    # Log current field values
                    Write-Log "Current field values:" -Level "DEBUG" -Color "Gray"
                    foreach ($field in $Fields.GetEnumerator()) {
                        $currentValue = $file.FieldValues[$field.Key]
                        Write-Log "  $($field.Key): '$currentValue' -> '$($field.Value)'" -Level "DEBUG" -Color "Gray"
                    }
                    
                    # Check out the file
                    Write-Log "→ Checking out file..." -Level "INFO" -Color "White"
                    Set-PnPFileCheckedOut -Url $fileUrl
                    Write-Log "✓ File checked out successfully" -Level "INFO" -Color "Green"
                    
                    # Update the metadata fields
                    Write-Log "→ Updating fields..." -Level "INFO" -Color "White"
                    Set-PnPListItem -List $Library -Identity $file.Id -Values $Fields
                    Write-Log "✓ Fields updated successfully" -Level "INFO" -Color "Green"
                    
                    # Check in the file with major version
                    Write-Log "→ Checking in file as major version..." -Level "INFO" -Color "White"
                    $checkinComment = "Manually approving with major version - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
                    Set-PnPFileCheckedIn -Url $fileUrl -CheckinType MajorCheckIn -Comment $checkinComment
                    Write-Log "✓ File checked in successfully with comment: '$checkinComment'" -Level "INFO" -Color "Green"
                    
                    $successCount++
                    Write-Log "✓ Successfully processed: $fileName" -Level "SUCCESS" -Color "Green"
                }
            }
            catch {
                Write-Log "✗ Error processing $fileName" -Level "ERROR" -Color "Red"
                Write-Log "  Error details: $($_.Exception.Message)" -Level "ERROR" -Color "Red"
                Write-Log "  Stack trace: $($_.ScriptStackTrace)" -Level "DEBUG" -Color "Gray"
                
                # Try to undo checkout if there was an error
                try {
                    Write-Log "→ Attempting to undo checkout..." -Level "WARN" -Color "Yellow"
                    if ([string]::IsNullOrEmpty($FolderPath)) {
                        $undoFileUrl = "$Library/$fileName"
                    } else {
                        $undoFileUrl = "$Library/$FolderPath/$fileName"
                    }
                    Set-PnPFileCheckedIn -Url $undoFileUrl -Comment "Error occurred - undoing checkout"
                    Write-Log "✓ Checkout undone successfully" -Level "WARN" -Color "Yellow"
                }
                catch {
                    Write-Log "✗ Could not undo checkout for $fileName: $($_.Exception.Message)" -Level "ERROR" -Color "Red"
                }
                
                $errorCount++
            }
        }
        
        # Summary
        Write-Log "" -Level "INFO" -Color "White"
        Write-Log "=== EXECUTION SUMMARY ===" -Level "INFO" -Color "Cyan"
        Write-Log "Successfully processed: $successCount files" -Level "SUCCESS" -Color "Green"
        Write-Log "Errors encountered: $errorCount files" -Level $(if($errorCount -gt 0) { "ERROR" } else { "INFO" }) -Color $(if($errorCount -gt 0) { "Red" } else { "Green" })
        Write-Log "Total files: $($Files.Count)" -Level "INFO" -Color "White"
        Write-Log "Execution completed at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level "INFO" -Color "White"
        
        if ($EnableLogging) {
            Write-Log "Detailed log saved to: $LogPath" -Level "INFO" -Color "Cyan"
        }
        
    }
    catch {
        Write-Log "Fatal Error: $($_.Exception.Message)" -Level "ERROR" -Color "Red"
        Write-Log "Could not connect to SharePoint or process files" -Level "ERROR" -Color "Red"
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level "DEBUG" -Color "Gray"
    }
    finally {
        # Disconnect from SharePoint
        try {
            Disconnect-PnPOnline
            Write-Log "Disconnected from SharePoint" -Level "INFO" -Color "Cyan"
        }
        catch {
            Write-Log "Could not disconnect properly: $($_.Exception.Message)" -Level "WARN" -Color "Yellow"
        }
    }
}

# Function to validate prerequisites
function Test-Prerequisites {
    # Check if PnP PowerShell module is installed
    if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
        Write-Host "PnP.PowerShell module not found!" -ForegroundColor Red
        Write-Host "Please install it using: Install-Module -Name PnP.PowerShell -Force" -ForegroundColor Yellow
        return $false
    }
    return $true
}

# Main execution
Write-Log "SharePoint File Update Script" -Level "INFO" -Color "Cyan"
Write-Log "=============================" -Level "INFO" -Color "Cyan"

# Initialize logging
Initialize-LogFile

if (Test-Prerequisites) {
    # Connect to SharePoint first
    Write-Log "Connecting to SharePoint site: $SiteUrl" -Level "INFO" -Color "Cyan"
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive
        Write-Log "Connected successfully!" -Level "INFO" -Color "Green"
    }
    catch {
        Write-Log "Failed to connect to SharePoint: $($_.Exception.Message)" -Level "ERROR" -Color "Red"
        exit
    }
    
    # Determine which files to process
    if ($AutoDiscoverFiles) {
        Write-Log "Auto-discovery mode enabled" -Level "INFO" -Color "White"
        # Auto-discover files with "In Progress" status
        $filesToProcess = Get-FilesInProgress -Library $DocumentLibrary -FolderPath $FolderPath
        
        if (Show-FilesForConfirmation -Files $filesToProcess) {
            $confirmation = Read-Host "`nDo you want to proceed with updating these files? (y/n)"
            Write-Log "User confirmation: $confirmation" -Level "INFO" -Color "Gray"
            
            if ($confirmation -eq 'y' -or $confirmation -eq 'Y') {
                Write-Log "User confirmed. Starting file updates..." -Level "INFO" -Color "Green"
                Update-SharePointFiles -SiteUrl $SiteUrl -Library $DocumentLibrary -Files $filesToProcess -Fields $FieldsToUpdate -UseFileObjects $true
            }
            else {
                Write-Log "Operation cancelled by user" -Level "WARN" -Color "Yellow"
            }
        }
    }
    else {
        Write-Log "Manual file specification mode" -Level "INFO" -Color "White"
        # Use manually specified files
        Write-Log "" -Level "INFO" -Color "White"
        Write-Log "Configuration:" -Level "INFO" -Color "White"
        Write-Log "Site URL: $SiteUrl" -Level "INFO" -Color "Gray"
        Write-Log "Library: $DocumentLibrary" -Level "INFO" -Color "Gray"
        if ([string]::IsNullOrEmpty($FolderPath)) {
            Write-Log "Folder: Root folder" -Level "INFO" -Color "Gray"
        } else {
            Write-Log "Folder: $FolderPath" -Level "INFO" -Color "Gray"
        }
        Write-Log "Files to update: $($FilesToUpdate.Count)" -Level "INFO" -Color "Gray"
        Write-Log "Columns to update:" -Level "INFO" -Color "Gray"
        foreach ($field in $FieldsToUpdate.GetEnumerator()) {
            Write-Log "  $($field.Key) = $($field.Value)" -Level "INFO" -Color "Gray"
        }
        
        $confirmation = Read-Host "`nDo you want to proceed? (y/n)"
        Write-Log "User confirmation: $confirmation" -Level "INFO" -Color "Gray"
        
        if ($confirmation -eq 'y' -or $confirmation -eq 'Y') {
            Write-Log "User confirmed. Starting file updates..." -Level "INFO" -Color "Green"
            Update-SharePointFiles -SiteUrl $SiteUrl -Library $DocumentLibrary -Files $FilesToUpdate -Fields $FieldsToUpdate -UseFileObjects $false
        }
        else {
            Write-Log "Operation cancelled by user" -Level "WARN" -Color "Yellow"
        }
    }
}
else {
    Write-Log "Prerequisites not met. Please install required modules." -Level "ERROR" -Color "Red"
}

# Final log message
if ($EnableLogging) {
    Write-Log "Script execution completed. Log file saved at: $LogPath" -Level "INFO" -Color "Cyan"
}
