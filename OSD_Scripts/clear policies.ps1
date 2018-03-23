    New-PSDrive -PSProvider Registry -Name Users -Root HKEY_USERS
    Set-Location Users:
     (Get-ChildItem "Users:\").PSPath | foreach { if($_){
            #write-host "Deleting " $_
            $policypath = $_ + "\software\policies"
            if (Test-Path -Path $policypath)
             {
             (Get-ChildItem $policypath) | foreach { if($_){
             Remove-Item $_ -Force -Recurse
             }}}
          } }

    $Folder = ("HKLM:\SOFTWARE\Policies\")

    #Write-Host "starting to delete predefined machine policies"
    #reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft" /f
    #reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google" /f
    #reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Hewlett-Packard" /f
    #reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\HP" /f
    
    #Write-Host "starting to delete machine Policies"
    (Get-ChildItem $Folder).PSPath | foreach { if($_){
            write-host "Deleting " $_
            Remove-Item $_ -Force -Recurse} }