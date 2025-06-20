# List of file extensions to add
$fileTypes = @(".py", ".ps1", ".csv", ".js", ".html", ".bat", ".md", ".json", ".xml")

foreach ($ext in $fileTypes) {
    $path = "HKLM:\Software\Classes\$ext\ShellNew"
    
    if (-not (Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
        New-ItemProperty -Path $path -Name "NullFile" -Value "" -PropertyType String | Out-Null
        Write-Host "Added $ext to context menu"
    } else {
        Write-Host "$ext already exists in context menu"
    }
}

Write-Host "Changes will take effect after restarting File Explorer or your computer"
