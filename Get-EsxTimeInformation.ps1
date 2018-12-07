$vCenterServer = Read-Host -Prompt "vCenter Server"
"Enter vCenter server credential"
$VIserverCredential = Get-Credential
"Enter ESX server credential"
$ESXCredential = Get-Credential

$ScriptName	= [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition)
$RepLog     = $PathScript
$FicLog     = $RepLog + $ScriptName + ".csv"

# Get all the ESX servers
Connect-VIserver -Server $vCenterServer -Credential $VIserverCredential
$VMhosts = Get-VMHost | Select-Object Name | Sort-Object -property Name
Disconnect-VIServer * -Confirm:$false

# Get the system time from all the ESX servers
$DateTimes = @{}
$VMHosts | ForEach-Object {
	Connect-VIServer -Server $_.Name -Credential $ESXCredential
	$DateTimes["$($_.Name)"] = (Get-View ServiceInstance).CurrentTime()
	Disconnect-VIServer * -Confirm:$false  
}

# Get the other information from the vCenter Server and combine it with the system time
Connect-VIserver -Server $vCenterServer -Credential $VIserverCredential
Get-VMHost | Sort-Object -property Name | ForEach-Object {
	$Report = "" | Select-Object -Property "ESX Name","Cluster Name","ESX Version",TimeZone,"Time Servers",DateTime,"ntpd Status"
	$ESX_Name = $_.Name
	$Cluster_Name = (Get-Cluster -VMHost $_).name
	$ESX_Version = $_.Version
	$TimeZone = $_.TimeZone
	$Time_Servers = "$(Get-VMHostNTPServer -VMHost $_)"
	$DateTime = $DateTimes["$($_.Name)"]
	If ((Get-VmHostService -VMHost $_ | Where-Object {$_.key -eq "ntpd"}).Running -eq "True") {
		$ntpd_status = "Running"	}
	Else {
		$ntpd_status = "Not Running"	}

	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject "$ESX_Name;$Cluster_Name;$ESX_Version;$TimeZone;$Time_Servers;$DateTime;$ntpd_status" -Append
}
Disconnect-VIServer * -Confirm:$False  