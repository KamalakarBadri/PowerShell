$TenantId = "0e439a1f-a497-462b-9e6b-4e582e203607"
$ClientId = "73efa35d-6188-42d4-b258-838a977eb149"
$ThumbPrint = "B799789F78628CAE56B4D0F380FD551EB754E0DB"

# Array of site URLs to process
$SiteUrls = @(
    "https://geekbyteonline.sharepoint.com/sites/New365",
    "https://geekbyteonline.sharepoint.com/sites/AnotherSite",
    "https://geekbyteonline.sharepoint.com/sites/ThirdSite"
)

$ExcludedLists = @("Access Requests", "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", 
    "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms",
    "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", 
    "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Images", "site collection images",
    "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", 
    "Reusable Content", "Reporting Metadata", "Reporting Templates", "Search Config List", "Site Assets", 
    "Preservation Hold Library", "Solution Gallery", "Style Library", "Suggested Content Browser Locations", 
    "Theme Gallery", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", 
    "Workflow History", "Workflow Tasks", "Pages")

# HTML Report Header with dynamic update capability
$htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>SharePoint Permissions Report</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 95%;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #0078d4;
            border-bottom: 3px solid #0078d4;
            padding-bottom: 10px;
        }
        .summary {
            background-color: #e8f4f8;
            padding: 15px;
            border-radius: 5px;
            margin: 20px 0;
            border-left: 4px solid #0078d4;
        }
        .summary h3 {
            margin-top: 0;
            color: #0078d4;
        }
        .site-card {
            background-color: white;
            border: 1px solid #ddd;
            border-radius: 5px;
            margin-bottom: 30px;
            overflow-x: auto;
        }
        .site-header {
            background-color: #0078d4;
            color: white;
            padding: 10px 15px;
            cursor: pointer;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .site-header:hover {
            background-color: #005a9e;
        }
        .site-title {
            font-size: 1.2em;
            font-weight: bold;
        }
        .site-stats {
            font-size: 0.9em;
            opacity: 0.9;
        }
        .toggle-icon {
            font-size: 1.2em;
        }
        .site-content {
            padding: 15px;
            display: none;
        }
        .site-content.active {
            display: block;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
            font-size: 0.85em;
        }
        th {
            background-color: #f2f2f2;
            padding: 12px;
            text-align: left;
            border: 1px solid #ddd;
            font-weight: bold;
            position: sticky;
            top: 0;
        }
        td {
            padding: 10px;
            border: 1px solid #ddd;
            vertical-align: top;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .permission-cell {
            max-width: 300px;
            word-wrap: break-word;
            white-space: pre-wrap;
            font-size: 0.8em;
        }
        .badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 3px;
            font-size: 0.75em;
            font-weight: bold;
        }
        .badge-read {
            background-color: #d1ecf1;
            color: #0c5460;
        }
        .badge-edit {
            background-color: #fff3cd;
            color: #856404;
        }
        .badge-full {
            background-color: #f8d7da;
            color: #721c24;
        }
        .badge-unique {
            background-color: #d4edda;
            color: #155724;
        }
        .loading {
            text-align: center;
            padding: 20px;
            color: #666;
        }
        .filter-bar {
            margin: 20px 0;
            padding: 10px;
            background-color: #f9f9f9;
            border-radius: 5px;
        }
        .filter-bar input {
            padding: 8px;
            margin: 5px;
            border: 1px solid #ddd;
            border-radius: 4px;
            width: 200px;
        }
        .filter-bar button {
            padding: 8px 15px;
            background-color: #0078d4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        .filter-bar button:hover {
            background-color: #005a9e;
        }
        .progress-container {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            background-color: #f0f0f0;
            z-index: 1000;
            display: none;
        }
        .progress-bar {
            width: 0%;
            height: 5px;
            background-color: #0078d4;
            transition: width 0.3s;
        }
        .progress-text {
            text-align: center;
            font-size: 0.8em;
            padding: 5px;
        }
        @media print {
            .site-header {
                print-color-adjust: exact;
                -webkit-print-color-adjust: exact;
            }
            .site-content {
                display: block !important;
            }
        }
    </style>
</head>
<body>
    <div class="progress-container" id="progressContainer">
        <div class="progress-bar" id="progressBar"></div>
        <div class="progress-text" id="progressText"></div>
    </div>
    <div class="container">
        <h1>📊 SharePoint Permissions Report</h1>
        <div class="summary" id="summarySection">
            <h3>Report Summary</h3>
            <div id="summaryContent">Loading report...</div>
        </div>
        <div class="filter-bar">
            <input type="text" id="searchInput" placeholder="Search by item name, library, or user..." onkeyup="filterTable()">
            <button onclick="expandAll()">Expand All</button>
            <button onclick="collapseAll()">Collapse All</button>
            <button onclick="exportToExcel()">Export to Excel</button>
        </div>
        <div id="reportContent">
            <div class="loading">Processing SharePoint sites... Please wait...</div>
        </div>
    </div>
    <script>
        function toggleSite(siteId) {
            var content = document.getElementById('content-' + siteId);
            var icon = document.getElementById('icon-' + siteId);
            if (content.classList.contains('active')) {
                content.classList.remove('active');
                icon.innerHTML = '▶';
            } else {
                content.classList.add('active');
                icon.innerHTML = '▼';
            }
        }
        
        function expandAll() {
            var contents = document.querySelectorAll('.site-content');
            var icons = document.querySelectorAll('.toggle-icon');
            contents.forEach(function(content) {
                content.classList.add('active');
            });
            icons.forEach(function(icon) {
                icon.innerHTML = '▼';
            });
        }
        
        function collapseAll() {
            var contents = document.querySelectorAll('.site-content');
            var icons = document.querySelectorAll('.toggle-icon');
            contents.forEach(function(content) {
                content.classList.remove('active');
            });
            icons.forEach(function(icon) {
                icon.innerHTML = '▶';
            });
        }
        
        function filterTable() {
            var input = document.getElementById('searchInput');
            var filter = input.value.toLowerCase();
            var tables = document.querySelectorAll('table');
            
            tables.forEach(function(table) {
                var rows = table.getElementsByTagName('tr');
                for (var i = 1; i < rows.length; i++) {
                    var row = rows[i];
                    var text = row.innerText.toLowerCase();
                    if (text.indexOf(filter) > -1) {
                        row.style.display = '';
                    } else {
                        row.style.display = 'none';
                    }
                }
            });
        }
        
        function exportToExcel() {
            var html = document.querySelector('.container').cloneNode(true);
            var tables = html.querySelectorAll('table');
            tables.forEach(function(table) {
                var rows = table.querySelectorAll('tr');
                rows.forEach(function(row) {
                    var cells = row.querySelectorAll('td');
                    cells.forEach(function(cell) {
                        var text = cell.innerText;
                        cell.innerHTML = text;
                    });
                });
            });
            
            var blob = new Blob([html.outerHTML], {type: 'application/vnd.ms-excel'});
            var link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'SharePoint_Permissions_Report.xls';
            link.click();
        }
        
        function updateProgress(current, total, siteName) {
            var container = document.getElementById('progressContainer');
            var bar = document.getElementById('progressBar');
            var text = document.getElementById('progressText');
            var percent = (current / total) * 100;
            container.style.display = 'block';
            bar.style.width = percent + '%';
            text.innerHTML = `Processing: \${siteName} - \${current} of \${total} items (\${Math.round(percent)}%)`;
            if (current === total) {
                setTimeout(function() {
                    container.style.display = 'none';
                }, 2000);
            }
        }
    </script>
</body>
</html>
"@

# Initialize HTML content
$htmlContent = $htmlHeader
$allReportData = @()

# Function to get user login names and display names
function Get-UserDetails {
    param($UserId, $SiteUrl)
    
    try {
        $userUrl = "$SiteUrl/_api/web/getuserbyid($UserId)"
        $user = Invoke-PnPSPRestMethod -Url $userUrl -Method Get -ErrorAction Stop
        return @{
            LoginName = $user.LoginName
            DisplayName = $user.Title
            Email = $user.Email
        }
    }
    catch {
        return @{
            LoginName = "Unknown"
            DisplayName = "Unknown"
            Email = ""
        }
    }
}

# Loop through each SharePoint site
$totalSites = $SiteUrls.Count
$currentSite = 0

foreach ($siteUrl in $SiteUrls) {
    $currentSite++
    Write-Host "`n====================================================================" -ForegroundColor Cyan
    Write-Host "CONNECTING TO SITE: $siteUrl ($currentSite of $totalSites)" -ForegroundColor Cyan
    Write-Host "====================================================================" -ForegroundColor Cyan
    
    try {
        Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId -ErrorAction Stop
        Write-Host "Successfully connected to $siteUrl" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to connect to $siteUrl : $_" -ForegroundColor Red
        continue
    }

    # Get all document libraries
    Write-Host "`nRetrieving document libraries..." -ForegroundColor Yellow
    try {
        $lists = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists" -Method Get -ErrorAction Stop
        Write-Host "Found $($lists.value.Count) lists in the site" -ForegroundColor Green
    }
    catch {
        Write-Host "Error retrieving lists: $_" -ForegroundColor Red
        continue
    }

    $siteReportData = @()
    $processedItems = 0
    $itemsWithPermissions = 0

    foreach ($list in $lists.value) {
        if ($list.BaseTemplate -ne 101) { 
            Write-Host "Skipping non-document library: $($list.Title) (BaseTemplate: $($list.BaseTemplate))" -ForegroundColor Gray
            continue 
        }

        if ($ExcludedLists -contains $list.Title) {
            Write-Host "Skipping excluded list: $($list.Title)" -ForegroundColor Gray
            continue
        }

        Write-Host "`nProcessing Document Library: $($list.Title)" -ForegroundColor Yellow

        $nextPageUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items?`$top=1000&`$expand=File,Folder"
        $pageCount = 0
        
        do {
            $pageCount++
            Write-Host "  Retrieving page $pageCount of items..." -ForegroundColor DarkGray
            
            try {
                $response = Invoke-PnPSPRestMethod -Url $nextPageUrl -Method Get -ErrorAction Stop
                $listItems = $response.value
                $nextPageUrl = $response."odata.nextLink"
                Write-Host "  Found $($listItems.Count) items on this page" -ForegroundColor DarkGray

                foreach ($item in $listItems) {
                    $processedItems++
                    
                    # Update progress in console
                    if ($processedItems % 50 -eq 0) {
                        Write-Host "Processed $processedItems items total (found $itemsWithPermissions with unique permissions so far)" -ForegroundColor DarkGray
                    }

                    try {
                        # Determine item type and get details
                        $itemName = $null
                        $itemLocation = $null
                        $itemSize = $null
                        $itemType = "Unknown"
                        
                        if ($item.FileSystemObjectType -eq 0 -and $item.File) {
                            $itemType = "File"
                            $itemName = $item.File.Name
                            $itemLocation = $item.File.ServerRelativeUrl
                            $itemSize = $item.File.Length
                            Write-Host "    Processing file: $itemName" -ForegroundColor DarkGray
                        } 
                        elseif ($item.FileSystemObjectType -eq 1 -and $item.Folder) {
                            $itemType = "Folder"
                            $itemName = $item.Folder.Name
                            $itemLocation = $item.Folder.ServerRelativeUrl
                            Write-Host "    Processing folder: $itemName" -ForegroundColor DarkGray
                        } 
                        else {
                            $itemType = "ListItem"
                            $itemName = $item.Title
                            $itemLocation = $null
                            Write-Host "    Processing list item: $itemName" -ForegroundColor DarkGray
                        }

                        # Check if item has unique permissions
                        try {
                            $uniquePerms = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/HasUniqueRoleAssignments" -Method Get -ErrorAction Stop
                            $hasUniquePerms = $uniquePerms.value
                        }
                        catch {
                            Write-Host "    Error checking unique permissions for item $($item.Id): $_" -ForegroundColor Red
                            $hasUniquePerms = $false
                        }

                        # Only get detailed permissions if item has unique permissions
                        $readUsers = @()
                        $editUsers = @()
                        $fullControlUsers = @()
                        $sharingLinks = @()

                        if ($hasUniquePerms) {
                            try {
                                # Get role assignments for this item
                                $roleAssignmentsUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/roleassignments?`$expand=Member,RoleDefinitionBindings"
                                $roleAssignments = Invoke-PnPSPRestMethod -Url $roleAssignmentsUrl -Method Get -ErrorAction Stop
                                
                                foreach ($ra in $roleAssignments.value) {
                                    $member = $ra.Member
                                    $roleBindings = $ra.RoleDefinitionBindings
                                    
                                    # Determine permission level
                                    $permissionLevel = ""
                                    foreach ($role in $roleBindings) {
                                        if ($role.Name -eq "Read" -or $role.Name -eq "Restricted Read") {
                                            $permissionLevel = "Read"
                                        }
                                        elseif ($role.Name -eq "Contribute" -or $role.Name -eq "Edit") {
                                            $permissionLevel = "Edit"
                                        }
                                        elseif ($role.Name -eq "Full Control" -or $role.Name -eq "Owner") {
                                            $permissionLevel = "FullControl"
                                        }
                                    }
                                    
                                    if ($member.PrincipalType -eq 1) {
                                        # User
                                        $userDetails = Get-UserDetails -UserId $member.Id -SiteUrl $siteUrl
                                        $userEntry = "$($userDetails.DisplayName) [$($userDetails.LoginName)]"
                                        
                                        switch ($permissionLevel) {
                                            "Read" { $readUsers += $userEntry }
                                            "Edit" { $editUsers += $userEntry }
                                            "FullControl" { $fullControlUsers += $userEntry }
                                        }
                                    }
                                    elseif ($member.PrincipalType -eq 4 -or $member.PrincipalType -eq 8) {
                                        # SharePoint Group
                                        $groupName = $member.Title
                                        $groupMembersUrl = "$siteUrl/_api/web/SiteGroups/GetById($($member.Id))/Users"
                                        try {
                                            $members = Invoke-PnPSPRestMethod -Url $groupMembersUrl -Method Get -ErrorAction Stop
                                            foreach ($groupMember in $members.value) {
                                                $userDetails = Get-UserDetails -UserId $groupMember.Id -SiteUrl $siteUrl
                                                $userEntry = "$($userDetails.DisplayName) [$($userDetails.LoginName)] (via $groupName)"
                                                
                                                switch ($permissionLevel) {
                                                    "Read" { $readUsers += $userEntry }
                                                    "Edit" { $editUsers += $userEntry }
                                                    "FullControl" { $fullControlUsers += $userEntry }
                                                }
                                            }
                                        }
                                        catch {
                                            $groupEntry = "$groupName [Group - members not accessible]"
                                            switch ($permissionLevel) {
                                                "Read" { $readUsers += $groupEntry }
                                                "Edit" { $editUsers += $groupEntry }
                                                "FullControl" { $fullControlUsers += $groupEntry }
                                            }
                                        }
                                    }
                                }
                                
                                # Check for sharing links (unique to this item)
                                try {
                                    $sharingUrl = "$siteUrl/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/GetSharingInformation"
                                    $sharingInfo = Invoke-PnPSPRestMethod -Url $sharingUrl -Method Get -ErrorAction Stop
                                    
                                    if ($sharingInfo.HasSharingLinks) {
                                        foreach ($link in $sharingInfo.SharingLinks) {
                                            $sharingLinks += "$($link.LinkUrl) ($($link.ShareType) access)"
                                        }
                                    }
                                }
                                catch {
                                    # No sharing links found
                                }
                            }
                            catch {
                                Write-Host "    Error getting role assignments: $_" -ForegroundColor Red
                            }
                        }

                        # Add to report only if there are permissions (unique or not)
                        $itemsWithPermissions++
                        $reportEntry = [PSCustomObject]@{
                            SiteName         = $siteUrl.Split("/")[-1]
                            SiteUrl          = $siteUrl
                            LibraryName      = $list.Title
                            ItemID           = $item.Id
                            ItemType         = $itemType
                            Name             = $itemName
                            Location         = $itemLocation
                            Size             = if ($itemSize) { "$([math]::Round($itemSize/1KB, 2)) KB" } else { "" }
                            ReadUsers        = ($readUsers | Sort-Object -Unique) -join "`n"
                            EditUsers        = ($editUsers | Sort-Object -Unique) -join "`n"
                            FullControlUsers = ($fullControlUsers | Sort-Object -Unique) -join "`n"
                            SharingLinks     = if ($sharingLinks.Count -gt 0) { ($sharingLinks | Sort-Object -Unique) -join "`n" } else { "None" }
                            UniquePerms      = if ($hasUniquePerms) { "Yes" } else { "No (inherited)" }
                            LastModified     = if ($item.Modified) { $item.Modified } else { "" }
                        }
                        $siteReportData += $reportEntry
                        $allReportData += $reportEntry
                        
                        if ($hasUniquePerms) {
                            Write-Host "      ✅ Found unique permissions for $itemType $itemName" -ForegroundColor Green
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

    # Generate HTML for this site
    if ($siteReportData.Count -gt 0) {
        $siteId = $siteUrl -replace '[^a-zA-Z0-9]', '_'
        $siteName = $siteUrl.Split("/")[-1]
        
        $siteHtml = @"
        <div class="site-card">
            <div class="site-header" onclick="toggleSite('$siteId')">
                <div>
                    <span class="site-title">📁 $siteName</span>
                    <div class="site-stats">
                        Items with permissions: $($siteReportData.Count) | 
                        Libraries: $($siteReportData.LibraryName | Select-Object -Unique | Measure-Object | Select-Object -ExpandProperty Count)
                    </div>
                </div>
                <span class="toggle-icon" id="icon-$siteId">▶</span>
            </div>
            <div class="site-content" id="content-$siteId">
                <table>
                    <thead>
                        <tr>
                            <th>Library</th>
                            <th>Item Type</th>
                            <th>Name</th>
                            <th>Location</th>
                            <th>Size</th>
                            <th>Read Users</th>
                            <th>Edit Users</th>
                            <th>Full Control Users</th>
                            <th>Sharing Links</th>
                            <th>Unique Permissions</th>
                            <th>Last Modified</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($item in $siteReportData) {
            $siteHtml += @"
                         <tr>
                            <td>$($item.LibraryName)</td>
                            <td>$($item.ItemType)</td>
                            <td>$($item.Name)</td>
                            <td>$($item.Location)</td>
                            <td>$($item.Size)</td>
                            <td class="permission-cell">$($item.ReadUsers -replace "`n", "<br/>")</td>
                            <td class="permission-cell">$($item.EditUsers -replace "`n", "<br/>")</td>
                            <td class="permission-cell">$($item.FullControlUsers -replace "`n", "<br/>")</td>
                            <td class="permission-cell">$($item.SharingLinks -replace "`n", "<br/>")</td>
                            <td>$($item.UniquePerms)</td>
                            <td>$($item.LastModified)</td>
                         </tr>
"@
        }
        
        $siteHtml += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
        
        # Add site HTML to report
        if ($htmlContent -match "(<div id=""reportContent"">)") {
            $htmlContent = $htmlContent -replace "(<div id=""reportContent"">)", "`$1$siteHtml"
        }
    }

    Disconnect-PnPOnline
    Write-Host "Disconnected from $siteUrl" -ForegroundColor DarkGray
}

# Update summary section
$totalItems = $allReportData.Count
$uniquePermsCount = ($allReportData | Where-Object { $_.UniquePerms -eq "Yes" }).Count
$totalLibraries = ($allReportData.LibraryName | Select-Object -Unique).Count
$totalSitesWithData = ($allReportData.SiteName | Select-Object -Unique).Count

$summaryHtml = @"
<p><strong>Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
<p><strong>Total Sites Processed:</strong> $totalSitesWithData</p>
<p><strong>Total Items with Permissions:</strong> $totalItems</p>
<p><strong>Items with Unique Permissions:</strong> $uniquePermsCount</p>
<p><strong>Total Libraries Processed:</strong> $totalLibraries</p>
<p><strong>Report Format:</strong> HTML with dynamic filtering and expand/collapse functionality</p>
"@

$htmlContent = $htmlContent -replace "(<div id=""summaryContent"">).*?(</div>)", "`$1$summaryHtml`$2"
$htmlContent = $htmlContent -replace '<div class="loading">.*?</div>', ''

# Save HTML report
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$htmlFileName = "SharePoint_Permissions_Report_$timestamp.html"
$htmlContent | Out-File -FilePath $htmlFileName -Encoding UTF8

Write-Host "`n====================================================================" -ForegroundColor Green
Write-Host "REPORT GENERATION COMPLETE" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Green
Write-Host "Total items processed: $($allReportData.Count)" -ForegroundColor White
Write-Host "HTML Report saved as: $htmlFileName" -ForegroundColor Green
Write-Host "Open the HTML file in your browser to view the report" -ForegroundColor Cyan
Write-Host "====================================================================" -ForegroundColor Green
