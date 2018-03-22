    New-PSDrive -PSProvider Registry -Name Users -Root HKEY_USERS
    #try
    #{
    write-host "User is NOT logged on"
    write-host "Stating to delete User Policies"
     (Get-ChildItem "Users:\").PSPath | foreach { if($_){
            write-host "Deleting " $_
            $policypath = $_ + "\software\policies"
            if (Test-Path -Path $policypath)
             {
             Write-Host $policypath
             (Get-ChildItem $policypath) | foreach { if($_){
             Remove-Item $_ -Force -Recurse
             }}
             write-host "test done"
             Write-Host "/n /n"
             }

            #Remove-Item $_ -Force -Recurse
          } }
    #}
    <#finally {
    Remove-PSDrive -Name HKUser

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    $retryCount = 0
    $retryLimit = 20
    $retryTime = 1 #seconds

    reg unload "HKU\$sid" # > $null

    while ($LASTEXITCODE -ne 0 -and $retryCount -lt $retryLimit) {
        Write-Verbose "Error unloading 'HKU\$sid', waiting and trying again." -Verbose
        Start-Sleep -Seconds $retryTime
        $retryCount++
        reg unload "HKU\$sid" 
    }
    }
    } 
    #>

    $Folder = ("HKLM:\SOFTWARE\Policies\")

    #Test-Path -Path $Folder 
    Write-Host "starting to delete predefined machine policies"
    reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft" /f
    reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google" /f
    reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Hewlett-Packard" /f
    reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\HP" /f
    
    
    
    Write-Host "starting to delete machine Policies"
    (Get-ChildItem $Folder).PSPath | foreach { if($_){
            write-host "Deleting " $_
            Remove-Item $_ -Force -Recurse} }
    