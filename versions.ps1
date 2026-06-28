# ============================================================
# SHAREPOINT FILE VERSION HISTORY REPORT GENERATOR
# COUNTS ALL VERSIONS INCLUDING CURRENT
# ============================================================

# ============================================================
# CONFIGURATION - UPDATE THESE VALUES
# ============================================================

$CONFIG = @{
    # Your SharePoint site URL
    "site_url" = "https://geekbyteonline.sharepoint.com/sites/Team_Site2"
    
    # ============================================================
    # FILE FILTER CONFIGURATION
    # ============================================================
    "file_extensions" =  $null  # Process ALL files
    "file_extensions" = @(".docx", ".xlsx", ".pdf")
    # ============================================================
    # VERSION HISTORY FILTER - CHECK ALL FILES
    # ============================================================
    "min_file_size_mb" = 0  # Check versions for ALL files
    
    "output_csv" = $null
}

# Global variables
$Script:CSVWriters = @{}
$Script:CSVFiles = @{}
$Script:LibrarySummary = @{}
$Script:PnPConnected = $false

# ============================================================
# PNP AUTHENTICATION FUNCTIONS
# ============================================================

function Install-PnPPowerShell {
    try {
        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
            Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
            Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
            Write-Host "PnP.PowerShell module installed successfully!" -ForegroundColor Green
        }
        Import-Module PnP.PowerShell -Force -ErrorAction Stop
        return $true
    }
    catch {
        Write-Host "Failed to install/import PnP.PowerShell module: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Connect-PnPSharePoint {
    param(
        [string]$SiteUrl
    )
    
    try {
        if (-not (Install-PnPPowerShell)) {
            throw "PnP.PowerShell module is required"
        }
        
        Write-Host "`n  🔐 Connecting to SharePoint: $SiteUrl" -ForegroundColor Cyan
        Write-Host "  Using -UseWebLogin" -ForegroundColor Cyan
        
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop
        
        $web = Get-PnPWeb -ErrorAction Stop
        Write-Host "  ✓ Connected to site: $($web.Title)" -ForegroundColor Green
        
        $Script:PnPConnected = $true
        return $true
    }
    catch {
        Write-Host "  ✗ Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Get-SiteUrl {
    $url = $CONFIG.site_url
    if ([string]::IsNullOrEmpty($url)) {
        throw "Site URL is not configured."
    }
    return $url.TrimEnd('/')
}

function Get-SiteName {
    param([string]$SiteUrl)

    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        return 'Site'
    }

    $cleanUrl = $SiteUrl.Trim().TrimEnd('/')

    if ($cleanUrl -match '/sites/([^/]+)') {
        return ($matches[1] -replace '[\\/:*?"<>|]', '_').Trim()
    }

    $segments = $cleanUrl.Split('/')
    if ($segments.Count -gt 0) {
        $lastSegment = $segments[-1]
        if ([string]::IsNullOrWhiteSpace($lastSegment) -or $lastSegment -eq 'sites') {
            return 'Site'
        }

        return ($lastSegment -replace '[\\/:*?"<>|]', '_').Trim()
    }

    return 'Site'
}

function Get-DefaultOutputFile {
    param([string]$SiteUrl)

    $siteName = Get-SiteName -SiteUrl $SiteUrl
    $timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    return "${siteName}_File_Version_History_Report_${timestamp}.csv"
}

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

function ConvertTo-MB {
    param([long]$Bytes)
    if ($Bytes -eq 0) { return 0.00 }
    return [math]::Round($Bytes / (1024 * 1024), 2)
}

function Format-DateTime {
    param([string]$DateTimeStr)
    if ([string]::IsNullOrEmpty($DateTimeStr) -or $DateTimeStr -eq "N/A") {
        return "N/A"
    }
    try {
        $dt = [datetime]::Parse($DateTimeStr)
        return $dt.ToString("yyyy-MM-dd HH:mm:ss")
    }
    catch {
        return $DateTimeStr
    }
}

function Test-ShouldProcessFile {
    param([string]$FileName)
    if ([string]::IsNullOrEmpty($FileName)) { return $false }
    
    $extConfig = $CONFIG.file_extensions
    if ($null -eq $extConfig) { return $true }
    
    if ($extConfig -is [array]) {
        if ($extConfig.Count -eq 0) { return $true }
        $fileNameLower = $FileName.ToLower()
        foreach ($ext in $extConfig) {
            if ($fileNameLower.EndsWith($ext.ToLower())) { return $true }
        }
        return $false
    }
    return $true
}

function Get-FilterDescription {
    $extConfig = $CONFIG.file_extensions
    if ($null -eq $extConfig) { return "All files (no extension filter)" }
    if ($extConfig -is [array]) {
        if ($extConfig.Count -eq 0) { return "All files (empty extension list)" }
        return "Only files with extensions: $($extConfig -join ', ')"
    }
    return "Unknown filter configuration"
}

# ============================================================
# CSV REPORT FUNCTIONS
# ============================================================

function Initialize-Reports {
    param([string]$OutputFile)
    
    if ([string]::IsNullOrEmpty($OutputFile)) {
        throw "Output file path is empty."
    }
    
    Write-Host "`nInitializing CSV report..." -ForegroundColor Cyan
    
    $mainFile = $OutputFile
    $Script:CSVFiles['main'] = [System.IO.StreamWriter]::new($mainFile, $false, [System.Text.Encoding]::UTF8)
    
    $headers = 'Library,List ID,Item ID,File Name,File Path,Current File Size (MB),Version Count,First Version Date,Last Version Date,Total Versions Size (MB),File Created,File Modified,Versions Checked,File Extension,Processed At'
    $Script:CSVFiles['main'].WriteLine($headers)
    $Script:CSVFiles['main'].Flush()
    Write-Host "✓ Main report initialized: $mainFile" -ForegroundColor Green
    
    Write-Host "  Filter: $(Get-FilterDescription)" -ForegroundColor Cyan
    Write-Host "  Version history will be checked for ALL files" -ForegroundColor Green
}

function Append-ToMainReport {
    param([hashtable]$Data)
    
    try {
        $row = @(
            $Data.library,
            $Data.library_id,
            $Data.item_id,
            $Data.file_name,
            $Data.file_path,
            ("{0:F2}" -f $Data.current_file_size_mb),
            $Data.version_count,
            $Data.first_version_formatted,
            $Data.last_version_formatted,
            ("{0:F2}" -f $Data.total_versions_size_mb),
            $Data.created_formatted,
            $Data.modified_formatted,
            $(if ($Data.versions_checked) { 'Yes' } else { 'No' }),
            $Data.file_extension,
            (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        )
        
        $quotedRow = $row | ForEach-Object {
            if ($_ -match '[,"]') {
                $escaped = $_ -replace '"', '""'
                "`"$escaped`""
            } else {
                $_
            }
        }
        
        $line = $quotedRow -join ','
        $Script:CSVFiles['main'].WriteLine($line)
        $Script:CSVFiles['main'].Flush()
        return $true
    }
    catch {
        return $false
    }
}

function Append-ToDetailedReport {
    param([hashtable]$Data)
    
    try {
        if ($Data.versions -and $Data.versions.Count -gt 0) {
            foreach ($version in $Data.versions) {
                $row = @(
                    $Data.library,
                    $Data.item_id,
                    $Data.file_name,
                    $Data.file_path,
                    $version.version_id,
                    $version.version_label,
                    $version.ui_version_string,
                    (Format-DateTime -DateTimeStr $version.created),
                    $(if ($version.is_current) { 'Yes' } else { 'No' }),
                    ("{0:F2}" -f (ConvertTo-MB -Bytes $version.size)),
                    $version.checkin_comment,
                    $version.author,
                    $version.editor,
                    (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                )
                
                $quotedRow = $row | ForEach-Object {
                    if ($_ -match '[,"]') {
                        $escaped = $_ -replace '"', '""'
                        "`"$escaped`""
                    } else {
                        $_
                    }
                }
                
                $line = $quotedRow -join ','
                $Script:CSVFiles['detailed'].WriteLine($line)
                $Script:CSVFiles['detailed'].Flush()
            }
        }
        return $true
    }
    catch {
        return $false
    }
}

function Update-SummaryReport {
    param([hashtable]$LibrarySummary)
    
    try {
        $summaryFile = $Script:CSVFiles['summary'].BaseStream.Name
        $Script:CSVFiles['summary'].Close()
        $Script:CSVFiles['summary'] = [System.IO.StreamWriter]::new($summaryFile, $false, [System.Text.Encoding]::UTF8)
        
        $headers = 'Library,Total Files,Files Checked,Files Skipped (Size),Files Skipped (Filter),Total Versions,Total Current Size (MB),Total Versions Size (MB),Average Versions per File,Last Updated'
        $Script:CSVFiles['summary'].WriteLine($headers)
        
        foreach ($lib in $LibrarySummary.Keys) {
            $stats = $LibrarySummary[$lib]
            $avgVersions = if ($stats.files_checked -gt 0) { 
                [math]::Round($stats.versions / $stats.files_checked, 2) 
            } else { 0 }
            
            $row = @(
                $lib,
                $stats.total_files,
                $stats.files_checked,
                $stats.files_skipped_size,
                $stats.files_skipped_filter,
                $stats.versions,
                ("{0:F2}" -f (ConvertTo-MB -Bytes $stats.current_size)),
                ("{0:F2}" -f (ConvertTo-MB -Bytes $stats.versions_size)),
                $avgVersions,
                (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            )
            
            $quotedRow = $row | ForEach-Object {
                if ($_ -match '[,"]') {
                    $escaped = $_ -replace '"', '""'
                    "`"$escaped`""
                } else {
                    $_
                }
            }
            
            $line = $quotedRow -join ','
            $Script:CSVFiles['summary'].WriteLine($line)
        }
        $Script:CSVFiles['summary'].Flush()
        return $true
    }
    catch {
        return $false
    }
}

function Close-Reports {
    foreach ($key in $Script:CSVFiles.Keys) {
        try {
            if ($Script:CSVFiles[$key]) {
                $Script:CSVFiles[$key].Flush()
                $Script:CSVFiles[$key].Close()
            }
        }
        catch { }
    }
}

# ============================================================
# SHAREPOINT DATA RETRIEVAL FUNCTIONS
# ============================================================

function Get-AllLibraries {
    param([string]$SiteUrl)
    
    Write-Host "`nGetting document libraries..." -ForegroundColor Cyan
    
    try {
        $lists = Get-PnPList -ErrorAction Stop
        $allLibraries = @()
        
        foreach ($lst in $lists) {
            if ($lst.BaseTemplate -eq 101) {
                $allLibraries += @{
                    id = $lst.Id
                    title = $lst.Title
                }
                Write-Host "  ✓ Found library: $($lst.Title)" -ForegroundColor Gray
            }
        }
        
        Write-Host "  Found $($allLibraries.Count) document libraries" -ForegroundColor Green
        return $allLibraries
    }
    catch {
        Write-Host "  ✗ Failed to get libraries: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

function Get-AllItemsFromLibrary {
    param(
        [string]$SiteUrl,
        [string]$LibraryId,
        [string]$LibraryTitle
    )
    
    Write-Host "    Fetching items from library: $LibraryTitle" -ForegroundColor Gray
    
    try {
        $list = Get-PnPList -Identity $LibraryId -ErrorAction Stop
        $items = Get-PnPListItem -List $list -PageSize 2000 -ErrorAction Stop
        
        Write-Host "    Total items fetched: $($items.Count)" -ForegroundColor Gray
        return $items
    }
    catch {
        Write-Host "    ✗ Failed to get items: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

function Get-FileVersions {
    param(
        [string]$SiteUrl,
        [string]$ListId,
        [int]$ItemId,
        [string]$FileName
    )
    
    try {
        # REST API call
        $restUrl = "$SiteUrl/_api/Web/Lists(guid'$ListId')/items($ItemId)/versions"
        
        $response = Invoke-PnPSPRestMethod -Url $restUrl -Method Get -ErrorAction Stop
        
        # Check if we got a response with the 'value' property
        if ($response -and $response.value) {
            $totalItems = $response.value.Count
            
            if ($totalItems -gt 0) {
                $versionList = @()
                
                # Process ALL versions (including current)
                foreach ($version in $response.value) {
                    $isCurrent = $version.IsCurrentVersion
                    $versionLabel = $version.VersionLabel
                    $uiVersion = $version.'OData__x005f_UIVersionString'
                    
                    $versionData = @{
                        version_id = $version.VersionId
                        version_label = $versionLabel
                        ui_version_string = $uiVersion
                        created = $version.Created
                        is_current = $isCurrent
                        size = [long]$version.'File_x005f_x0020_x005f_Size'
                        checkin_comment = $version.'OData__x005f_CheckinComment'
                        author = ''
                        editor = ''
                    }
                    
                    if ($version.Author -and $version.Author.LookupValue) {
                        $versionData.author = $version.Author.LookupValue
                    } elseif ($version.'Created_x005f_x0020_x005f_By') {
                        $versionData.author = $version.'Created_x005f_x0020_x005f_By'
                    }
                    
                    if ($version.Editor -and $version.Editor.LookupValue) {
                        $versionData.editor = $version.Editor.LookupValue
                    } elseif ($version.'Modified_x005f_x0020_x005f_By') {
                        $versionData.editor = $version.'Modified_x005f_x0020_x005f_By'
                    }
                    
                    $versionList += $versionData
                }
                
                # Return ALL versions (including current)
                return $versionList
            }
        }
        
        return @()
    }
    catch {
        return @()
    }
}

function Get-FileDetailsFromPnPItem {
    param([PSObject]$Item)
    
    $fileName = $Item.FieldValues.FileLeafRef
    if ([string]::IsNullOrEmpty($fileName)) {
        $fileName = $Item.FieldValues.Title
    }
    if ([string]::IsNullOrEmpty($fileName)) {
        $fileName = "Item_$($Item.Id)"
    }
    
    $filePath = $Item.FieldValues.FileRef
    if ([string]::IsNullOrEmpty($filePath)) {
        $filePath = $Item.FieldValues.FileDirRef + "/" + $fileName
    }
    
    $fileSize = [long]$Item.FieldValues.File_x0020_Size
    if ($fileSize -eq 0) {
        $fileSize = [long]$Item.FieldValues.SMTotalSize
    }
    
    $fileExtension = ''
    if ($fileName) {
        $fileExtension = [System.IO.Path]::GetExtension($fileName).ToLower()
    }
    
    $created = if ($Item.FieldValues.Created) { $Item.FieldValues.Created } else { 'N/A' }
    $modified = if ($Item.FieldValues.Modified) { $Item.FieldValues.Modified } else { 'N/A' }
    
    return @{
        file_name = $fileName
        file_path = $filePath
        file_size = $fileSize
        file_extension = $fileExtension
        created = $created
        modified = $modified
    }
}

function Process-FileItem {
    param(
        [string]$SiteUrl,
        [string]$ListId,
        [PSObject]$Item,
        [string]$LibraryTitle,
        [hashtable]$LibrarySummary
    )
    
    try {
        $itemId = $Item.Id
        
        # Check if it's a folder
        if ($Item.FieldValues.FileSystemObjectType -eq "Folder") {
            return $null
        }
        
        $fileDetails = Get-FileDetailsFromPnPItem -Item $Item
        $fileSizeMB = ConvertTo-MB -Bytes $fileDetails.file_size
        $fileExtension = $fileDetails.file_extension
        
        # Check if file should be processed based on extension filter
        if (-not (Test-ShouldProcessFile -FileName $fileDetails.file_name)) {
            return $null
        }
        
        # IMPORTANT: Initialize $versions as an empty array BEFORE getting versions
        $versions = @()
        
        # Get versions - this returns ALL versions including current
        $versions = Get-FileVersions -SiteUrl $SiteUrl -ListId $ListId -ItemId $itemId -FileName $fileDetails.file_name
        
        # Ensure $versions is an array
        if ($versions -isnot [array]) {
            $versions = @()
        }
        
        # Initialize version count
        $versionCount = 0
        $totalVersionsSize = 0
        $firstVersionDate = 'N/A'
        $lastVersionDate = 'N/A'
        
        # Count ALL versions including current
        if ($versions -and $versions.Count -gt 0) {
            $versionCount = $versions.Count
            
            # Sort versions by creation date
            $sortedVersions = $versions | Sort-Object -Property created
            $firstVersion = $sortedVersions[0]
            $lastVersion = $sortedVersions[-1]
            
            $firstVersionDate = $firstVersion.created
            $lastVersionDate = $lastVersion.created
            
            # Calculate total size of all versions
            foreach ($version in $versions) {
                $totalVersionsSize += $version.size
            }
        } else {
            # No versions found, use current file size
            $totalVersionsSize = $fileDetails.file_size
            $versionCount = 0
        }
        
        $fileData = @{
            library = $LibraryTitle
            library_id = $ListId
            item_id = $itemId
            file_name = $fileDetails.file_name
            file_path = $fileDetails.file_path
            file_extension = $fileExtension
            current_file_size = $fileDetails.file_size
            current_file_size_mb = $fileSizeMB
            version_count = $versionCount
            first_version_date = $firstVersionDate
            last_version_date = $lastVersionDate
            total_versions_size = $totalVersionsSize
            total_versions_size_mb = ConvertTo-MB -Bytes $totalVersionsSize
            versions = $versions
            versions_checked = $true
            created_formatted = Format-DateTime -DateTimeStr $fileDetails.created
            modified_formatted = Format-DateTime -DateTimeStr $fileDetails.modified
            first_version_formatted = Format-DateTime -DateTimeStr $firstVersionDate
            last_version_formatted = Format-DateTime -DateTimeStr $lastVersionDate
        }
        
        # Update library summary
        if (-not $LibrarySummary.ContainsKey($LibraryTitle)) {
            $LibrarySummary[$LibraryTitle] = @{
                total_files = 0
                files_checked = 0
                files_skipped_size = 0
                files_skipped_filter = 0
                versions = 0
                current_size = 0
                versions_size = 0
            }
        }
        
        $LibrarySummary[$LibraryTitle].total_files++
        $LibrarySummary[$LibraryTitle].files_checked++
        $LibrarySummary[$LibraryTitle].versions += $versionCount
        $LibrarySummary[$LibraryTitle].current_size += $fileDetails.file_size
        $LibrarySummary[$LibraryTitle].versions_size += $totalVersionsSize
        
        # Append to main report
        Append-ToMainReport -Data $fileData
        
        return $fileData
    }
    catch {
        Write-Host "    ✗ Error processing item $($Item.Id): $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# ============================================================
# MAIN PROCESSING FUNCTIONS
# ============================================================

function Process-Files {
    param(
        [string]$SiteUrl,
        [string]$OutputFile
    )
    
    Initialize-Reports -OutputFile $OutputFile
    
    $libraries = Get-AllLibraries -SiteUrl $SiteUrl
    
    if (-not $libraries -or $libraries.Count -eq 0) {
        Write-Host "No document libraries found." -ForegroundColor Red
        Close-Reports
        return @()
    }
    
    Write-Host "`nFound $($libraries.Count) document libraries:" -ForegroundColor Green
    foreach ($lib in $libraries) {
        Write-Host "  - $($lib.title)" -ForegroundColor Gray
    }
    
    Write-Host "`nFilter: $(Get-FilterDescription)" -ForegroundColor Cyan
    Write-Host "Version history will be checked for ALL files" -ForegroundColor Green
    
    $allFileData = @()
    $totalFiles = 0
    $processed = 0
    $Script:LibrarySummary = @{}
    
    foreach ($library in $libraries) {
        Write-Host "`n$('='*60)" -ForegroundColor Cyan
        Write-Host "Processing library: $($library.title)" -ForegroundColor Yellow
        Write-Host "$('='*60)" -ForegroundColor Cyan
        
        $items = Get-AllItemsFromLibrary -SiteUrl $SiteUrl -LibraryId $library.id -LibraryTitle $library.title
        
        if (-not $items -or $items.Count -eq 0) {
            Write-Host "  No items found in $($library.title)" -ForegroundColor Yellow
            continue
        }
        
        # Filter out folders
        $files = @()
        foreach ($item in $items) {
            if ($item.FieldValues.FileSystemObjectType -ne "Folder") {
                $files += $item
            }
        }
        
        if (-not $files -or $files.Count -eq 0) {
            Write-Host "  No files found in $($library.title)" -ForegroundColor Yellow
            continue
        }
        
        # Filter files based on extension filter
        $filteredFiles = @()
        $filterSkipped = 0
        
        foreach ($f in $files) {
            $fileName = $f.FieldValues.FileLeafRef
            if (Test-ShouldProcessFile -FileName $fileName) {
                $filteredFiles += $f
            }
            else {
                $filterSkipped++
            }
        }
        
        if (-not $filteredFiles -or $filteredFiles.Count -eq 0) {
            Write-Host "  No files match the filter criteria in $($library.title)" -ForegroundColor Yellow
            Write-Host "  ($filterSkipped files skipped due to filter)" -ForegroundColor Gray
            continue
        }
        
        Write-Host "  Found $($filteredFiles.Count) files in $($library.title) (after filter)" -ForegroundColor Green
        if ($filterSkipped -gt 0) {
            Write-Host "  - $filterSkipped files skipped due to extension filter" -ForegroundColor Gray
        }
        Write-Host "  - ALL files will have versions checked" -ForegroundColor Green
        
        $totalFiles += $filteredFiles.Count
        
        foreach ($fileItem in $filteredFiles) {
            $processed++
            $itemId = $fileItem.Id
            $fileName = $fileItem.FieldValues.FileLeafRef
            if ([string]::IsNullOrEmpty($fileName)) {
                $fileName = "Item_$itemId"
            }
            $fileSizeMB = ConvertTo-MB -Bytes $fileItem.FieldValues.File_x0020_Size
            
            Write-Host "`n  [$processed/$totalFiles] Processing: $fileName (ID: $itemId) [$($fileSizeMB.ToString('F2')) MB]" -NoNewline
            
            $fileData = Process-FileItem -SiteUrl $SiteUrl -ListId $library.id -Item $fileItem -LibraryTitle $library.title -LibrarySummary $Script:LibrarySummary
            
            if ($fileData) {
                $allFileData += $fileData
                if ($fileData.version_count -gt 0) {
                    Write-Host " ✓ ($($fileData.version_count) versions, $($fileData.total_versions_size_mb.ToString('F2')) MB total)" -ForegroundColor Green
                }
                else {
                    Write-Host " ✓ (0 versions found)" -ForegroundColor Gray
                }
            }
            else {
                Write-Host " ✗ (Failed to process)" -ForegroundColor Red
            }
            
            Start-Sleep -Milliseconds 300
        }
    }
    
    Write-Host "`n$('='*60)" -ForegroundColor Cyan
    Write-Host "Processed $($allFileData.Count) files." -ForegroundColor Green
    
    Close-Reports
    
    return $allFileData
}

# ============================================================
# MAIN FUNCTION
# ============================================================

function Main {
    Write-Host "$('='*80)" -ForegroundColor Cyan
    Write-Host "FILE VERSION HISTORY REPORT GENERATOR" -ForegroundColor Yellow
    Write-Host "(COUNTS ALL VERSIONS INCLUDING CURRENT)" -ForegroundColor Yellow
    Write-Host "$('='*80)" -ForegroundColor Cyan
    $siteUrl = Get-SiteUrl
    $siteName = Get-SiteName -SiteUrl $siteUrl
    $outputFile = if ([string]::IsNullOrWhiteSpace($CONFIG.output_csv)) { Get-DefaultOutputFile -SiteUrl $siteUrl } else { $CONFIG.output_csv }
    $CONFIG.output_csv = $outputFile

    Write-Host "SharePoint Site: $siteUrl" -ForegroundColor White
    Write-Host "Site Name: $siteName" -ForegroundColor White
    Write-Host "Filter: $(Get-FilterDescription)" -ForegroundColor White
    Write-Host "Version Check: ALL files (no size limit)" -ForegroundColor Green
    Write-Host "Output File: $outputFile" -ForegroundColor White
    Write-Host "$('='*80)" -ForegroundColor Cyan
    
    # Connect to SharePoint
    Write-Host "`nAuthenticating to SharePoint..." -ForegroundColor Cyan
    $connected = Connect-PnPSharePoint -SiteUrl $CONFIG.site_url
    
    if (-not $connected) {
        Write-Host "✗ Failed to connect to SharePoint." -ForegroundColor Red
        return
    }
    
    Write-Host "✓ Connection established`n" -ForegroundColor Green
    
    Write-Host "Starting file processing..." -ForegroundColor Cyan
    Write-Host "Using REST API for version retrieval" -ForegroundColor Gray
    
    $startTime = Get-Date
    $fileData = Process-Files -SiteUrl $siteUrl -OutputFile $outputFile
    $elapsedTime = (Get-Date) - $startTime
    
    Write-Host "`nProcessing completed in $($elapsedTime.TotalSeconds.ToString('F2')) seconds." -ForegroundColor Green
    
    if (-not $fileData -or $fileData.Count -eq 0) {
        Write-Host "`nNo files were processed." -ForegroundColor Yellow
        return
    }
    
    # Print final summary
    Write-Host "`n$('='*80)" -ForegroundColor Cyan
    Write-Host "PROCESSING COMPLETED SUCCESSFULLY!" -ForegroundColor Green
    Write-Host "$('='*80)" -ForegroundColor Cyan
    
    $filesWithVersions = ($fileData | Where-Object { $_.version_count -gt 0 }).Count
    $filesWithoutVersions = $fileData.Count - $filesWithVersions
    $totalVersions = ($fileData | Measure-Object -Property version_count -Sum).Sum
    
    Write-Host "Total files processed: $($fileData.Count)" -ForegroundColor White
    Write-Host "Files with versions: $filesWithVersions" -ForegroundColor Green
    Write-Host "Files without versions: $filesWithoutVersions" -ForegroundColor Gray
    Write-Host "Total versions found: $totalVersions" -ForegroundColor White
    
    Write-Host "$('='*80)" -ForegroundColor Cyan
    Write-Host "✓ Main Report: $outputFile" -ForegroundColor Green
    Write-Host "$('='*80)" -ForegroundColor Cyan
}

# ============================================================
# SCRIPT EXECUTION
# ============================================================

if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "This script requires PowerShell 5.0 or higher." -ForegroundColor Red
    exit 1
}

Main