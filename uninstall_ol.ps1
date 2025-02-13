# Run with highest privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    try {
        Start-Process PowerShell -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -WindowStyle Hidden
        exit
    }
    catch {
        exit
    }
}

try {
    Set-ExecutionPolicy Bypass -Scope Process -Force
    $ErrorActionPreference = 'SilentlyContinue'

    # Close Outlook processes
    Get-Process | Where-Object { $_.ProcessName -like "*outlook*" } | Stop-Process -Force
    Start-Sleep -Seconds 2

    # Remove Outlook apps
    Get-AppxPackage *Microsoft.Office.Outlook* | Remove-AppxPackage
    Get-AppxProvisionedPackage -Online | Where-Object {$_.PackageName -like "*Microsoft.Office.Outlook*"} | Remove-AppxProvisionedPackage -Online
    Get-AppxPackage *Microsoft.OutlookForWindows* | Remove-AppxPackage
    Get-AppxProvisionedPackage -Online | Where-Object {$_.PackageName -like "*Microsoft.OutlookForWindows*"} | Remove-AppxProvisionedPackage -Online

    # Remove Outlook folders
    $windowsAppsPath = "C:\Program Files\WindowsApps"
    $outlookFolders = Get-ChildItem -Path $windowsAppsPath -Directory | Where-Object { $_.Name -like "Microsoft.OutlookForWindows*" }
    foreach ($folder in $outlookFolders) {
        $folderPath = Join-Path $windowsAppsPath $folder.Name
        takeown /f $folderPath /r /d Y | Out-Null
        icacls $folderPath /grant administrators:F /t | Out-Null
        Remove-Item -Path $folderPath -Recurse -Force
    }

    # Remove shortcuts
    $shortcutPaths = @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Outlook.lnk",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Outlook.lnk",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Outlook.lnk",
        "$env:PUBLIC\Desktop\Outlook.lnk",
        "$env:USERPROFILE\Desktop\Outlook.lnk",
        "$env:PUBLIC\Desktop\Microsoft Outlook.lnk",
        "$env:USERPROFILE\Desktop\Microsoft Outlook.lnk",
        "$env:PUBLIC\Desktop\Outlook (New).lnk",
        "$env:USERPROFILE\Desktop\Outlook (New).lnk",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook (New).lnk",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Outlook (New).lnk"
    )
    $shortcutPaths | ForEach-Object { Remove-Item $_ -Force }

    # Taskbar cleanup
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -Value 0 -Type DWord -Force

    $registryPaths = @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\TaskbarMRU",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\TaskBar",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    )
    foreach ($path in $registryPaths) {
        if (Test-Path $path) {
            @("Favorites", "FavoritesResolve", "FavoritesChanges", "FavoritesRemovedChanges", "TaskbarWinXP", "PinnedItems") | 
            ForEach-Object { Remove-ItemProperty -Path $path -Name $_ -ErrorAction SilentlyContinue }
        }
    }

    Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\Shell\LayoutModification.xml" -Force
    Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\iconcache*" -Force
    Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\thumbcache*" -Force

    # Restart Explorer
    Get-Process explorer | Stop-Process -Force
    Start-Sleep -Seconds 2
    Start-Process explorer
}
catch {}
exit 