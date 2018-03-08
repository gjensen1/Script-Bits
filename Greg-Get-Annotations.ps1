function greg-get-annotations {
<#
.DESCRIPTION
Greg-get-annotations function stores information about annotation fields for vms in given
cluster or in all clusters in VC. It stores the result in an arraylist $vms, you can either create
a csv report from this object or display it on screen
greg-get-annotations |export-csv -NoTypeInformation c:\file1.csv will export it to csv file etc...
greg-get-annotations |format-table VMname,Cluster,CreatedOn,Notes will just display on screen a table with
annotations that include : vm name, its cluster and field "CreatedOn" and Notes
 
.PARAMETER clustername
Specifies the clustername against wchi report will be built
 
.EXAMPLE
greg-get-annotations -clustername 'cluster01'|Export-Csv c:\annotation-report.csv
Will procude report on vms that resides in 'cluster01' and store it in csv file
 
.EXAMPLE
greg-get-annotations -clustername 'cluster01'|ft *
Will procude report on vms that resides in 'cluster01' output it to screen
 
.EXAMPLE
greg-get-annotations |Export-Csv c:\annotation-report.csv
Will procude report on vms that resides in all clusters and output it to screen
 
.EXAMPLE
greg-get-annotations
Without specified -clustername switch, it will do report regarding all clusters in VC
 
.NOTES
AUTHOR: Grzegorz Kulikowski
LASTEDIT: 05/30/2011
 
 
#>
param ([string]$clustername)
    if(!($clustername)){$clusters=Get-Cluster}else{$clusters=Get-Cluster $clustername}
    $VMs=New-Object Collections.ArrayList
    foreach ($cluster in $clusters)  {
        foreach ($vmview in (get-view -ViewType VirtualMachine -SearchRoot $cluster.id)) {
            $vm=New-Object PsObject
            $vm 
            Add-Member -InputObject $vm -MemberType NoteProperty -Name VMname -Value $vmview.Name
            Add-Member -InputObject $vm -MemberType NoteProperty -Name Notes -Value $vmview.Config.Annotation
            Add-Member -InputObject $vm -MemberType NoteProperty -Name Cluster -Value $cluster.Name
            foreach ($CustomAttribute in $vmview.AvailableField){
                Add-Member -InputObject $vm -MemberType NoteProperty -Name $CustomAttribute.Name -Value ($vmview.Summary.CustomValue | ? {$_.Key -eq $CustomAttribute.Key}).value
            }
            $VMs.add($vm)|Out-Null
        }
    }
#return $VMs
}