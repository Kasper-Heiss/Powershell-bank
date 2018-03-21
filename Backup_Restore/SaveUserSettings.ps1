Param($Path)

[System.Xml.XmlDocument]$Settings = New-Object System.Xml.XmlDocument
[System.Xml.XmlElement]$Root = $Settings.CreateElement("root")
$Settings.AppendChild($Root) | Out-Null

$LanguageCodes = $null

Function ReadLanguages() {
    Write-Host "Read Language Data ... " -ForegroundColor Green -NoNewline
    $LanguageData = $null
    If (Test-Path "$PSScriptRoot\Languages.csv") {
        $LanguageData = Import-Csv "$PSScriptRoot\Languages.csv" -Delimiter ';'
    }
    Write-Host "Done" -ForegroundColor Green
    return $LanguageData
}

Function AddElement($Name, $Value = $null, [System.Xml.XmlElement]$rootNode) {

    [System.Xml.XmlElement]$NewElement = $Settings.CreateElement($name)

    if ($value) {
        $EncodedValue = [Security.SecurityElement]::Escape($Value)

        $NewTextNode = $Settings.CreateTextNode($EncodedValue)
        $NewElement.AppendChild($NewTextNode) 
    }
    $temp = $rootNode.AppendChild($NewElement) 

    return $NewElement
}

Function AddCData($Name, $Value, [System.Xml.XmlElement]$rootNode) {

    $el = AddElement -Name "$Name" -rootNode $rootNode
    $result = $Settings.CreateCDataSection($Value)
    $temp = $el.AppendChild($result) 

    return $el
}

Function TranslateProofingTools($ID) {

    $Language = $LanguageCodes | Where { $_.Hex -eq "0x$ID" }

    If ($Language) {
        Return $Language.Language
    } else {
        Return "Unknown"
    }
}

Function TranslateKeyboardLayouts($ID) {

    $ID = $ID.substring(4)

    $Language = $LanguageCodes | Where { $_.Hex -eq "0x$ID" }

    If ($Language) {
        Return $Language.Tag
    } else {
        Return "Unknown"
    }
}

Function SaveFavoriteFiles() {

    Write-Host "Gather Favorites ... " -ForegroundColor Green -NoNewline

    $FavNode = AddElement -Name "Favorites" -rootNode $Root

    Add-Type -Assembly "System.IO.Compression.FileSystem"

    $FavoritesFolder = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" | Select -ExpandProperty Favorites

    $TempFileName = [System.IO.Path]::GetTempFileName()

    If (Test-Path $TempFileName) {
        Remove-Item $TempFileName 
    }

    [System.io.compression.zipfile]::CreateFromDirectory($FavoritesFolder, $TempFileName)

    $content = [System.IO.File]::ReadAllBytes($TempFileName)
    $contentEncoded = [System.Convert]::ToBase64String($content)

    AddCData -Name "IEFavorites" -Value $contentEncoded -rootNode $FavNode | Out-Null

    Write-Host "Done" -ForegroundColor Green
}

Function SaveChromeFavoriteFiles() {

    Write-Host "Gather ChromeFavorites ... " -ForegroundColor Green -NoNewline

    $FavNode = AddElement -Name "ChromeFavorites" -rootNode $Root

    Add-Type -Assembly "System.IO.Compression.FileSystem"

    $FavoritesFolder = $env:LOCALAPPDATA + "\Google\Chrome\User Data\Default\Bookmarks" 

    $TempFileName = [System.IO.Path]::GetTempFileName()

    If (Test-Path $TempFileName) {
        Remove-Item $TempFileName 
    }

    $test = Get-Content -Path $FavoritesFolder
     
   #[System.io.compression.zipfile]::CreateFromDirectory($FavoritesFolder, $TempFileName)

    $content = [System.IO.File]::ReadAllBytes($FavoritesFolder)

    $contentEncoded = [System.Convert]::ToBase64String($content)

    AddCData -Name "Favorites" -Value $contentEncoded -rootNode $FavNode 

    Write-Host "Done" -ForegroundColor Green
}

Function SaveNetworkPrinters() {

    Write-Host "Gather Network Printers ... " -ForegroundColor Green -NoNewline

    $PrintersNode = AddElement -Name "Printers" -rootNode $Root

    $Printers = Get-WmiObject "Win32_Printer" -Filter "Network = True"

    Foreach ($Printer in $Printers) {

        $CurrentElement = AddElement -Name "Printer" -rootNode $PrintersNode

        AddElement -Name "Name" -Value $Printer.Name -rootNode $CurrentElement | Out-Null
        AddElement -Name "ShareName" -Value $printer.ShareName -rootNode $CurrentElement | Out-Null

    }
    Write-Host "Done" -ForegroundColor Green
}

Function SaveNetworkDrives() {
    Write-Host "Gather Network Drives ... " -ForegroundColor Green -NoNewline

    $DrivesNode = AddElement -Name "NetworkDrives" -rootNode $Root

    $Drives = Get-WmiObject "Win32_NetworkConnection" -Filter "Persistent = True"

    Foreach ($Drive in $Drives) {

        If ($Drive.UserName -ieq "$($Env:USERDOMAIN)\$($env:USERNAME)") {
            $CurrentElement = AddElement -Name "Drive" -rootNode $DrivesNode

            AddElement -Name "LocalName" -Value $Drive.LocalName -rootNode $CurrentElement | Out-Null
            AddElement -Name "RemotePath" -Value $Drive.RemotePath -rootNode $CurrentElement | Out-Null
        }
    }
    Write-Host "Done" -ForegroundColor Green
}

Function SaveOutlookSignatures() {

    Write-Host "Gather Outlook Signatures ... " -ForegroundColor Green -NoNewline

    $SignaturesNode = AddElement -Name "Signatures" -rootNode $Root

    Add-Type -Assembly "System.IO.Compression.FileSystem"

    $AppDataFolder = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" | Select -ExpandProperty AppData
    $SignatureFolder = "$AppDataFolder\Microsoft\Signatures"

    If (Test-Path $SignatureFolder) {
        $TempFileName = [System.IO.Path]::GetTempFileName()

        If (Test-Path $TempFileName) {
            Remove-Item $TempFileName 
        }

        [System.io.compression.zipfile]::CreateFromDirectory($SignatureFolder, $TempFileName)

        $content = [System.IO.File]::ReadAllBytes($TempFileName)
        $contentEncoded = [System.Convert]::ToBase64String($content)

        AddCData -Name "OutlookSignatures" -Value $contentEncoded -rootNode $SignaturesNode | Out-Null
    }
    Write-Host "Done" -ForegroundColor Green
}

Function SaveOfficeLanguageSettings() {
    
    Write-Host "Gather Office Language Settings ... " -ForegroundColor Green -NoNewline

    $OfficeNode = AddElement -Name "OfficeLanguage" -rootNode $Root

    If (Test-Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\LanguageResources") {
        $Value = Get-ItemProperty -ErrorAction SilentlyContinue -path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\LanguageResources" | Select -ExpandProperty UIFallbackLanguages -ErrorAction SilentlyContinue
        AddElement -Name "UIFallbackLanguages" -Value $Value -rootNode $OfficeNode | Out-Null
        $ProofingTools = Get-childitem -recurse -path hklm:\software\wow6432node\microsoft\windows\currentversion\Uninstall | foreach { if ($_.PSPath -match "{90160000-001F-*") { $_.PSChildName.SubString(15,4) } }
        $ProofingNode = AddElement -Name "ProofingTools" -rootNode $OfficeNode

        forEach($Tool in $ProofingTools) {
            #$Value = TranslateProofingTools -ID $Tool
            $Value = $Tool

            AddElement -Name "Language" -Value $Value -rootNode $ProofingNode | Out-Null
        }
    }

   

    Write-Host "Done" -ForegroundColor Green
}

Function SaveWindowsSettings() {

    Write-Host "Gather Windows Settings ... " -ForegroundColor Green -NoNewline

    $WindowsNode = AddElement -Name "WindowsSettings" -rootNode $Root

    $DisplayLanguage = Get-Host | Select -ExpandProperty CurrentUICulture | Select -ExpandProperty Name

    AddElement -Name "DisplayLanguage" -Value $DisplayLanguage -rootNode $WindowsNode | Out-Null

    add-type -AssemblyName System.Windows.Forms
    
    $KeyboardSettings = AddElement -Name "KeyboardSettings" -rootNode $WindowsNode 

    $CurrentKeyboard = [System.Windows.Forms.InputLanguage]::CurrentInputLanguage.Culture.Name

    AddElement -Name "CurrentLayout" -Value $CurrentKeyboard -rootNode $KeyboardSettings | Out-Null

    $KeyboardLayouts = Get-Item -Path "HKCU:\Keyboard Layout\Preload\" | select -ExpandProperty Property

    $LayoutNode = AddElement -Name "Layouts" -rootNode $KeyboardSettings 

    foreach ($layout in $KeyboardLayouts) {
        $LanguageCode = Get-ItemPropertyValue -Path "HKCU:\Keyboard Layout\Preload\" -Name $layout

            $LanguageName = TranslateKeyboardLayouts($LanguageCode)

           AddElement -Name "Language" -Value $LanguageName -rootNode $LayoutNode | Out-Null

    }

    $GEOLocation = Get-ItemPropertyValue -path "HKCU:\Control Panel\International\Geo" -Name "Nation"

    AddElement -Name "GEOLocation" -Value $GEOLocation -rootNode $WindowsNode | Out-Null
    
    $ActivePowerScheme = Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\Default\PowerSchemes" -Name "ActivePowerScheme"

    AddElement -Name "ActivePowerScheme" -Value $ActivePowerScheme -rootNode $WindowsNode | Out-Null

    Write-Host "Done" -ForegroundColor Green
}




## MAIN

# PS Version check
If ($PSVersionTable.PSVersion.Major -lt 5) {
    write-host "This script requires PowerShell version 5" -ForegroundColor Red
    Exit 1
}

Write-Host "Check requirements ... " -NoNewline -ForegroundColor Green

# Generate Name
$NowDate = Get-Date -f "dd-MM-yyyy"
$ExportFileName = "$($env:USERNAME)##$($env:COMPUTERNAME)##$($NowDate).xml"

If (!$Path) {
    #$Path = $PSScriptRoot
    $Path = "\\danfoss.net\files\Common\Settings\$env:USERNAME\"
    
    }

If (!(Test-Path $Path)) {
    #Create folder
    New-Item $Path -type directory -force | Out-Null
         
    $objACL = Get-ACL $Path
    
    #Set Modify Rights for DANFOSS\CS-DS_Administrator on the folder
    $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule("DANFOSS\CS-DS_Administrator", "Modify","None","None","Allow") 
    $objACL.AddAccessRule($objACE)
    Set-ACL $Path $objACL
    
    #Set Modify Rights for Current User on the folder and block inheritance 
    $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule($env:USERNAME, "Modify","None","None","Allow") 
    $objACL.AddAccessRule($objACE)
    $objACL.SetAccessRuleProtection($True, $False)    #Block Inheritance
    Set-ACL $Path $objACL
        
}

$CachedPath = "$($env:LOCALAPPDATA)\Settings"
If (!(Test-Path $CachedPath)) {
    #Create folder
    New-Item $CachedPath -type directory -force | Out-Null
}


Write-Host "Done" -ForegroundColor Green

$LanguageCodes = ReadLanguages

SaveNetworkDrives
SaveNetworkPrinters
SaveOfficeLanguageSettings
SaveWindowsSettings
SaveOutlookSignatures
SaveFavoriteFiles
SaveChromeFavoriteFiles

Write-Host "Save settings to file ... " -ForegroundColor Green -NoNewline
$Settings.Save("$Path\$ExportFileName")
$Settings.Save("$CachedPath\$ExportFileName")
Write-Host "Done" -ForegroundColor Green


# Clean all previously stored files, except the latest X (specified by limit)
$limit = 7
Get-ChildItem -Path $path -Recurse -Force | Where-Object { !$_.PSIsContainer} | sort CreationTime -descending | select -Skip $limit | Remove-Item -Force
Get-ChildItem -Path $CachedPath -Recurse -Force | Where-Object { !$_.PSIsContainer} | sort CreationTime -descending | select -Skip $limit | Remove-Item -Force

