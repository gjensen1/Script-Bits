Connect-VIServer 142.145.180.9 -Username root -Password "password"

$vms = get-vm | where { $_.PowerState -eq "PoweredOn" }

foreach ( $vm in $vms ) {Shutdown-VMGuest -VM $vm -Confirm:$false }

sleep 60

Restart-VMHost -Force -Confirm:$false