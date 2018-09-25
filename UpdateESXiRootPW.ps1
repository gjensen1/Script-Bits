#$ESXiServers     = 'esx01.domain','esx02.domain'
#$CurrentPassword = 'MyFakePassword'
#$NewPassword     = 'F@keNrTwo'


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

#***********************
# Function Get-NewRootPW
#***********************
Function Get-NewRootPW {
    [CmdletBinding()]
    Param()
    #Prompt User for ESXi Host Root Password
    $NewRootPW = Read-Host -assecurestring "Please enter the New Root password for the ESXi Hosts"
    $NewRootPW = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewRootPw))
    Return $NewRootPW
}
#**************************
# EndFunction Get-NewRootPW
#**************************

#**********************
# Function Get-FileName
#**********************
Function Get-FileName {
    [CmdletBinding()]
    Param($initialDirectory)
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "TXT (*.txt)| *.txt"
    $OpenFileDialog.ShowDialog() | Out-Null
    Return $OpenFileDialog.filename
}
#*************************
# EndFunction Get-FileName
#*************************

#*************************
# Function Read-TargetList
#*************************
Function Read-TargetList {
    [CmdletBinding()]
    Param($TargetFile)
    $Targets = Get-Content $TargetFile
    Return $Targets
}
#****************************
# EndFunction Read-TargetList
#****************************

"Get Target List"
$inputFile = Get-FileName $Global:WorkingFolder
#$inputFile = "$Global:WorkingFolder\Targets.txt"
"Reading Target List"
$ESXiServers = Read-TargetList $inputFile

$CurrentPassword = Get-RootPW
$NewPassword = Get-NewRootPW

$ESXiServers | ForEach-Object {
  try {
    Connect-VIServer $_ -User root -Password $CurrentPassword 
    Set-VMHostAccount -UserAccount root -Password $NewPassword
  } catch {
    throw $_
  } finally {
    Disconnect-VIServer -Confirm:$False -ea 0
  }
}