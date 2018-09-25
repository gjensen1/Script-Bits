[CmdletBinding(DefaultParameterSetName='DefaultConfiguration')]
Param
(        
    [Parameter(Mandatory=$true)][String]$Location,
    [Parameter(Mandatory=$true)][String]$DPMServername,

    [Switch]$CustomizeDPMSubscriptionSettings,
    [Switch]$SetEncryption,
    [Switch]$SetProxy
)

DynamicParam
{
    $paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
    $attributes = New-Object System.Management.Automation.ParameterAttribute
    $attributes.ParameterSetName = "__AllParameterSets"
    $attributes.Mandatory = $true
    $attributeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
    $attributeCollection.Add($attributes)
    # If "-SetEncryption" is used, then add the "EncryptionPassPhrase" parameter
    if($SetEncryption)
    { 
        $dynParam1 = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("EncryptionPassPhrase", [String], $attributeCollection)   
        $paramDictionary.Add("EncryptionPassPhrase", $dynParam1)
    }
    # If "-SetProxy" is used, then add the "ProxyServerAddress" "ProxyServerPort" and parameters
    if($SetProxy)
    {
        $dynParam1 = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ProxyServerAddress", [String], $attributeCollection)   
        $paramDictionary.Add("ProxyServerAddress", $dynParam1)
        $dynParam2 = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ProxyServerPort", [String], $attributeCollection)   
        $paramDictionary.Add("ProxyServerPort", $dynParam2)
    }
    # If "-CustomizeDPMSubscriptionSettings" is used, then add the "StagingAreaPath" parameter
    if($CustomizeDPMSubscriptionSettings)
    {
        $dynParam1 = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("StagingAreaPath", [String], $attributeCollection)   
        $paramDictionary.Add("StagingAreaPath", $dynParam1)
    }
    return $paramDictionary
}
Process{
    foreach($key in $PSBoundParameters.keys)
    {
        Set-Variable -Name $key -Value $PSBoundParameters."$key" -Scope 0
    }
}