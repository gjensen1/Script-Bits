﻿# -----------------------
# Define Global Variables
# -----------------------
$Global:Folder = $env:USERPROFILE+"\Documents\"

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
$inputFile = Get-FileName $Global:Folder
"----------------------------------------------------------"
"Reading Target List"
$TargetList = Read-TargetList $inputFile
"----------------------------------------------------------"

Invoke-Command -ComputerName $TargetList {
#Invoke-Command -ComputerName itspajweemra001 {
    $Patches = 'KB4088875', 'KB4088878'
    #$Patches = 'KB4074587'
    Get-HotFix -Id $Patches
} -ErrorAction SilentlyContinue -ErrorVariable Problem
#-Credential (Get-Credential) -ErrorAction SilentlyContinue -ErrorVariable Problem
 
foreach ($p in $Problem) {
    if ($p.origininfo.pscomputername) {
        Write-Warning -Message "Patch not found on $($p.origininfo.pscomputername)" 
    }
    elseif ($p.targetobject) {
        Write-Warning -Message "Unable to connect to $($p.targetobject)"
    }
}