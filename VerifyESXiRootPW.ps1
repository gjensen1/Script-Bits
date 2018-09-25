param(
    [Parameter(Mandatory=$true)][String]$vCenter
    )


#********************
# Function Get-RootPW
#********************
Function Get-RootPW {
    [CmdletBinding()]
    Param()
    #Prompt User for ESXi Host Root Password
    $RootPW = Read-Host -assecurestring "Please enter the Root password for the ESXi Hosts"
    $RootPW = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($RootPw))
    Return $RootPW
}
#***********************
# EndFunction Get-RootPW
#***********************

#*******************
# Connect to vCenter
#*******************
Function Connect-VC {
    [CmdletBinding()]
    Param()
    "Connecting to $Global:VCName"
    Connect-VIServer $Global:VCName -Credential $Global:Creds -WarningAction SilentlyContinue
    #Connect-VIServer $Global:VCName -WarningAction SilentlyContinue
}
#***********************
# EndFunction Connect-VC
#***********************
$Global:VCName = $vCenter
Connect-VC
$RootPW = Get-RootPW
get-vmhost | %{$null = connect-viserver $_.name -user root -password $RootPW -EA 0; if (-not ($?)) {write-warning "Password failed for $($_.name)"  } else {Disconnect-VIServer $_.name -force -confirm:$false} }


