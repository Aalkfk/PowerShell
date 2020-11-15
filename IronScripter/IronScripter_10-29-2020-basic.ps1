Function Get-RegisteredSettings {
    $RegOwner = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\' RegisteredOwner).RegisteredOwner
    $RegOrg = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\' RegisteredOrganization).RegisteredOrganization

    $RegisteredSettings = New-Object psobject
    $RegisteredSettings | Add-Member NoteProperty 'Registered Owner' $RegOwner
    $RegisteredSettings | Add-Member NoteProperty 'Registered Organization' $RegOrg
    $RegisteredSettings
}

Function Set-RegisteredSettings {
    param(
        [string]$RegOwner,
        [string]$RegOrg
    )
    If($RegOwner -eq $null -and $RegOrg -eq $null){
        Write-Host 'You have to select at least one option. Please try again! :-)'
    }
    else{
        #Set for x64 OS
        If([System.Environment]::Is64BitOperatingSystem -eq $True){
            If($RegOwner -ne $null -and $RegOrg -eq $null){
                Write-Host "$ENV:COMPUTERNAME - Registered Owner (x64, x86) was set to $RegOwner."
            }
            If($RegOrg -ne $null -and $RegOwner -eq $null){
                Write-Host "$ENV:COMPUTERNAME - Registered Organization (x64, x86) was set to $RegOrg."
            }
            If($RegOwner -ne $null -and $RegOrg -ne $null){
                Write-Host "$ENV:COMPUTERNAME - Registered Owner (x64, x86) was set to $RegOwner and Registered Organization (x64, x86) was set to $RegOrg."
            }
        }
        #Set for x86 OS
        else{
            If($RegOwner -ne $null -and $RegOrg -eq $null){
                Write-Host "$ENV:COMPUTERNAME - Registered Owner  was set to $RegOwner."
            }
            If($RegOrg -ne $null -and $RegOwner -eq $null){
                Write-Host "$ENV:COMPUTERNAME - Registered Organization was set to $RegOrg."
            }
            If($RegOwner -ne $null -and $RegOrg -ne $null){
                Write-Host "$ENV:COMPUTERNAME - Registered Owner was set to $RegOwner and Registered Organization was set to $RegOrg."
            }
        }                  
    }
}