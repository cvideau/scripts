$PathScript = ($myinvocation.MyCommand.Definition | Split-Path -parent) + "\"
$user       = "CSP_SCRIPT_ADM"
$fickey     = "Z:\Scripts\Credentials\key.crd"
$ficcred    = "Z:\Scripts\Credentials\vmware_adm.crd"
$key        = get-content $fickey
$pwd        = Get-Content $ficcred | ConvertTo-SecureString -key $key
$Credential = New-Object System.Management.Automation.PSCredential $user, $pwd

### >>>>>>>>>> SOUS-FONCTIONS <<<<<<<<<<<<
Function LogTrace ($Message){
	$Message = (Get-Date -format G) + " " + $Message
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Message -Append }

### >>>>>>> DEBUT DU SCRIPT <<<<<<<<<<<
#$ErrorActionPreference = "SilentlyContinue"
$rc			= 0
$ScriptName	= [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition)
$vbcrlf		= "`r`n"
$dat		= Get-Date -UFormat "%Y_%m_%d_%H_%M_%S"

### Création du répertoire de LOG si besoin
$RepLog     = $PathScript + "LOG"
If (!(Test-Path $RepLog)) { New-Item -Path $PathScript -ItemType Directory -Name "LOG"}
$RepLog     = $RepLog + "\"
$FicLog     = $RepLog + $ScriptName + "_" + $dat + ".csv"

### Si le fichier LOG n'existe pas on le crée à vide
If (!(Test-Path $FicLog)) {
	$Line = (Get-Date -format G) + " " + $Line
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Line
} Else {
	LogTrace ($Line)
}

# Setup array with hosts 
$vCenter = "SWMUZV1VCSZY.zres.ztech"

$rccnx = Connect-VIServer -Server $vcenter -Protocol https -Credential $Credential
	
$topCnxVcc = "0"
If ($rccnx -ne $null) {	If ($rccnx.Isconnected) { $topCnxVcc = "1" } }

If ($topCnxVcc -ne "1") {
	Write-Host "ERREUR: Connexion KO au vCenter $vCenter => Arrêt du script" -ForegroundColor White -BackgroundColor Red
	$rc += 1
	Exit $rc }
Else { Write-Host "SUCCES: Connexion OK au vCenter $vCenter`r`n" -ForegroundColor Black -BackgroundColor Green	}

$ESXCredential = Get-Credential
$oDatacenters = Get-Datacenter -Server $vCenter
ForEach($DC in $oDatacenters){
	$Clusters = Get-Cluster -Location $DC -Server $vCenter
	ForEach($Cluster in $Clusters){
		$ESXs = Get-vmHost -Location $Cluster -Server $vCenter
		ForEach($ESX in $ESXs) {
			Write-Host $ESX
			$Config = Get-VMHost $ESX -Server $vCenter | Select Name, Version, TimeZone, ConnectionState, PowerState
			$ESX_Version = $Config.Version
			$TimeZone = $Config.TimeZone
			$State = $Config.ConnectionState
			$Power = $Config.PowerState
			
			If ((Get-VmHostService -VMHost $ESX  -Server $vCenter | Where-Object {$_.key -eq "ntpd"}).Running -eq "True") {
				$ntpd_status = "Running"	}
			Else {
				$ntpd_status = "Not Running"	}
			
			$Gateway = Get-VmHostNetwork -Host $ESX -Server $vCenter | Select VMkernelGateway -ExpandProperty VirtualNic | Where {$_.VMotionEnabled} | Select -ExpandProperty VMkernelGateway
			$NTPServer = Get-VMHostNtpServer -VMHost $ESX -Server $vCenter
			
			$CurrentDateTime = (Get-Date)
			Connect-VIServer -Server $ESX -Credential $ESXCredential
			$DateTime = (Get-View ServiceInstance).CurrentTime()

			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject "$ESX;$vCenter;$Cluster;$ESX_Version;$TimeZone;$Gateway;$NTPServer;$CurrentDateTime;$DateTime;$ntpd_status;$State;$Power" -Append
		}
	}
}