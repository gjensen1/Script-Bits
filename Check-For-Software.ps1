#**************************
# Function Check-PowerCLI10 
#**************************
Function Check-PowerCLI10 {
    [CmdletBinding()]
    Param()
    #Check for Prereqs for the script
    #This includes, PowerCLI 10, plink, and pscp

    #Check for PowerCLI 10
    $powercli = Get-Module -ListAvailable VMware.PowerCLI
    if (!($powercli.version.Major -eq "10")) {
        Throw "VMware PowerCLI 10 is not installed on your system!!!"
    }
    Else {
        Write-Host "PowerCLI 10 is Installed" -ForegroundColor Green
    } 
}
#*****************************
# EndFunction Check-PowerCLI10
#*****************************

#*********************
# Function Check-Putty 
#*********************
Function Check-Putty {
    $Putty = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where DisplayName -Like "PuTTY*"
    If (!($Putty)){
        Throw "Putty is not installed on your system!!!"
    }
    Else {
        $PuttyName = $Putty.DisplayName
        Write-Host "$PuttyName is installed" -ForegroundColor Green
    }

}
#************************
# EndFunction Check-Putty
#************************

#**********************
# Function Check-WinSCP 
#**********************
Function Check-WinSCP {
    $WinSCP = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where DisplayName -Like "WinSCP*"
    If (!($WinSCP)){
        Throw "WinSCP is not installed on your system!!!"
    }
    Else {
        $WinSCPName = $WinSCP.DisplayName
        Write-Host "$WinSCPName is installed" -ForegroundColor Green
    }
}
#*************************
# EndFunction Check-WinSCP
#*************************

Check-PowerCLI10
Check-Putty
Check-WinSCP



Write-Host "still running"