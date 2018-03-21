Param(
    $Path,
    $userName,
    [switch]$Favorites,
    [switch]$NetworkPrinters,
    [switch]$NetworkDrives,
    [switch]$OutlookSignatures,
    [switch]$OfficeLanguageSettings,
    [switch]$WindowsSettings
)
##Start of test with GIT
[System.Xml.XmlDocument]$Settings = New-Object System.Xml.XmlDocument

Function ReadSettings() {
    Write-Host "Read settings from file ... " -NoNewline -ForegroundColor Green

    $result = $false

    if (!$Path) {
        #If a path is not specified then test if local cache store exists 

        $CachedPath = "$($env:LOCALAPPDATA)\Settings\"
        If ((Test-Path $CachedPath)) {
            $Path = $CachedPath }
        Else {
            # Locall cache store not found - use network store
            $Path = "\\danfoss.net\files\Common\Settings\$($env:USERNAME)\"
        }
    }
    
    

    If (!$userName) {
        $userName = $env:username
    }

    $files = Get-ChildItem -Path $Path -Filter "$userName##$env:COMPUTERNAME.xml"

    if (!$files) {
        $files = Get-ChildItem -Path $Path -Filter "$userName##*.xml" | sort -Property LastModifiedDate | Select -First 1
    }

    if ($files) {
        $filename = $files[0].FullName
        $Settings.Load($Filename)
        $result = $true
    }
    
    Write-Host "Done" -ForegroundColor Green
    Return $result
}

Function RestoreFavoriteFiles() {
    Write-Host "Restore Favorites ... " -NoNewline -ForegroundColor Green

    $FavoritesFolder = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" | Select -ExpandProperty Favorites

    Add-Type -Assembly "System.IO.Compression.FileSystem"

    $TempFileName = [System.IO.Path]::GetTempFileName()

    If (Test-Path $TempFileName) { 
        Remove-Item  $TempFileName
    }

    $Data = $Settings.Root.Favorites.IEFavorites.InnerText
    $Bytes = [System.Convert]::FromBase64String($data)
    [System.IO.File]::WriteAllBytes($TempFileName, $Bytes)

    $TempDirName = "$([System.IO.Path]::GetTempPath())$([System.IO.Path]::GetRandomFileName())"

    [System.IO.Compression.ZipFile]::ExtractToDirectory($TempFileName, $TempDirName)

    # Copy items from tempdir to fav folder
    Get-ChildItem -Path $TempDirName | copy-item -Destination $FavoritesFolder -Force:$false -ErrorAction SilentlyContinue -Container -Recurse | Out-Null
    Remove-Item -Path $TempDirName -Recurse -Force
    Remove-Item -Path $TempFileName -Force

    Write-Host "Done" -ForegroundColor Green
}

Function RestoreChromeFavoriteFiles() {
    Write-Host "Restore ChromeFavorites ... " -NoNewline -ForegroundColor Green

    $FavoritesFolder = $env:LOCALAPPDATA + "\Google\Chrome\User Data\Default\Bookmarks"

    Add-Type -Assembly "System.IO.Compression.FileSystem"

    $TempFileName = [System.IO.Path]::GetTempFileName()

    If (Test-Path $TempFileName) { 
        Remove-Item  $TempFileName
    }

    $Data = $Settings.Root.ChromeFavorites.Favorites.InnerText
    write-host "DATA START"
    write-host $Data
    write-host "DATA END"
    $Bytes = [System.Convert]::FromBase64String($data)
    [System.IO.File]::WriteAllBytes($TempFileName, $Bytes)

    $TempDirName = "$([System.IO.Path]::GetTempPath())$([System.IO.Path]::GetRandomFileName())"

    #[System.IO.Compression.ZipFile]::ExtractToDirectory($TempFileName, $TempDirName)

    
    # Copy items from tempdir to fav folder
    #Get-ChildItem -Path $TempDirName
    Write-Host $TempFileName
    copy-item -Path $TempFileName -Destination $FavoritesFolder -Force:$true #-ErrorAction SilentlyContinue -Container #| Out-Null
    #Remove-Item -Path $TempDirName -Recurse -Force
    #Remove-Item -Path $TempFileName -Force

    Write-Host "Done" -ForegroundColor Green
}

Function RestoreNetworkPrinters() {
    Write-Host "Restore Network Printers ... " -NoNewline -ForegroundColor Green

    Foreach ($printer in $Settings.root.Printers.ChildNodes) {
        Add-Printer -ConnectionName $printer.Name
    }

    Write-Host "Done" -ForegroundColor Green
}

Function RestoreNetworkDrives() {
    Write-Host "Restore Network Drives ... " -NoNewline -ForegroundColor Green

    Foreach ($drive in $Settings.root.NetworkDrives.ChildNodes) {
        $DriveLetter = $drive.LocalName -replace ":", ""
        New-PSDrive -Name "$DriveLetter" -PSProvider FileSystem -Root "$($drive.RemotePath)" -Scope Global -Persist -ErrorAction SilentlyContinue | Out-Null
    }

    Write-Host "Done" -ForegroundColor Green
}

Function RestoreOutlookSignatures() {
    Write-Host "Restore Outlook Signatures ... " -NoNewline -ForegroundColor Green

    $AppDataFolder = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" | Select -ExpandProperty AppData
    $SignatureFolder = "$AppDataFolder\Microsoft\Signatures"

    Add-Type -Assembly "System.IO.Compression.FileSystem"

    $TempFileName = [System.IO.Path]::GetTempFileName()

    If (Test-Path $TempFileName) { 
        Remove-Item  $TempFileName
    }

    $Data = $Settings.Root.Signatures.OutlookSignatures.InnerText

    $Bytes = [System.Convert]::FromBase64String($data)
    [System.IO.File]::WriteAllBytes($TempFileName, $Bytes)

    $TempDirName = "$([System.IO.Path]::GetTempPath())$([System.IO.Path]::GetRandomFileName())"

    [System.IO.Compression.ZipFile]::ExtractToDirectory($TempFileName, $TempDirName)

    # Copy items from tempdir to fav folder
    Get-ChildItem -Path $TempDirName -Recurse | copy-item -Destination $SignatureFolder -Force
    Remove-Item -Path $TempDirName -Recurse -Force
    Remove-Item -Path $TempFileName -Force

    Write-Host "Done" -ForegroundColor Green
}

Function RestoreOfficeLanguageSettings() {
    Write-Host "Restore Office Language Settings ... " -NoNewline -ForegroundColor Green

    ### To be decided

    $RegistryPath = "HKCU:\Software\Danfoss\Office\ProofingTools"

    New-Item -Path $RegistryPath -Force | Out-Null

    Foreach ($Tool in $Settings.root.OfficeLanguage.ProofingTools.Language ) {
        New-ItemProperty -Path $RegistryPath -Name $Tool -PropertyType String | Out-Null
    }
    
    Write-Host "Done" -ForegroundColor Green
}

Function RestoreWindowsSettings() {
    Write-Host "Restore Windows Settings ... " -NoNewline -ForegroundColor Green

    $DisplayLanguage = $Settings.root.WindowsSettings.DisplayLanguage

    $CurrentKeyboard = $Settings.root.WindowsSettings.KeyboardSettings.CurrentLayout
    $AllKeyboardLayouts = $Settings.root.WindowsSettings.KeyboardSettings.Layouts
    $Layouts = @()

    $Layouts += $CurrentKeyboard

    $Layouts += $AllKeyboardLayouts.Language | Where { $_ -ine $CurrentKeyboard }

    Set-WinUserLanguageList -LanguageList $Layouts -Force

    $GEOLocation = $Settings.root.WindowsSettings.GEOLocation
    Set-WinHomeLocation -GeoId $GEOLocation

    $ActivePowerScheme = $Settings.root.WindowsSettings.ActivePowerScheme
    PowerCFG.exe /S $ActivePowerScheme

    Write-Host "Done" -ForegroundColor Green
}

## MAIN

# PS Version check
If ($PSVersionTable.PSVersion.Major -lt 5) {
    write-host "This script requires PowerShell version 5" -ForegroundColor Red
    Exit 1
}



if (ReadSettings) {

    # Default to restore all settings
    $RestoreAll = $true
    RestoreChromeFavoriteFiles
    # If any switch is specified, do not restore all
    Foreach ($param in $PSBoundParameters.GetEnumerator()) {
        If ($param.Value -eq $true) {
            $RestoreAll = $false
        }
    }
    
    
    <#if ($RestoreAll -or $NetworkDrives ) { RestoreNetworkDrives }
    if ($RestoreAll -or $NetworkPrinters ) { RestoreNetworkPrinters }
    if ($RestoreAll -or $Favorites) { RestoreFavoriteFiles }
    if ($RestoreAll -or $OutlookSignatures ) { RestoreOutlookSignatures }
    if ($RestoreAll -or $OfficeLanguageSettings ) { RestoreOfficeLanguageSettings }
    if ($RestoreAll -or $WindowsSettings ) { RestoreWindowsSettings }#>
} else{
    Write-Host "No matching settings where found" -ForegroundColor Red
}
#>