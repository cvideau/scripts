Add-PSsnapin VMware.VimAutomation.Core

$vCenters			= $args[0] # Nom du vCenter
$DCexc				= $args[1] # Datacenter exclus
$DCinc				= $args[2] # Datacenter inclus
$ClustExc			= $args[3] # Clusters exclus
$ClustInc			= $args[4] # Clusters inclus
$ESXexc				= $args[5] # ESX exclus
$ESXinc				= $args[6] # ESX inclus
$TagExc				= $args[7] # TAGs exclus
$TagInc				= $args[8] # TAGs inclus
$Action				= $args[9] # Action à réaliser

# Bouchon pour TESTS
#$vCenters	= "swmuzv1vcszc.zres.ztech"
#$ClustExc	= "AUCUN"
#$DCexc		= "AUCUN"
#$DCinc		= "TOUS"
#$ESXexc	= "AUCUN"
#$ClustInc	= "CL_MU_GSA_Z2"
#$ESXinc	= "TOUS"
#$TagInc	= "Z_MU_ESX_Production"
#$TagExc	= "AUCUN"
#$Action	= "GET"

If (($vCenters -eq $NULL) -OR ($DCexc -eq $NULL) -OR ($DCinc -eq $NULL) -OR ($ClustExc -eq $NULL) -OR ($ClustInc -eq $NULL) -OR ($ESXexc -eq $NULL) -OR ($ESXinc -eq $NULL) -OR ($TagExc -eq $NULL) -OR ($TagInc -eq $NULL) -OR ($Action -eq $NULL)) { Write-Host "L'une des paramètres de commande à une valeur vide..." $host.Exit() }


$PathScript = ($myinvocation.MyCommand.Definition | Split-Path -parent) + "\"
$user       = "CSP_SCRIPT_ADM"
$fickey     = "D:\Scripts\Credentials\key.crd "
$ficcred    = "D:\Scripts\Credentials\vmware_adm.crd"
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
$FicLog     = $RepLog + $ScriptName + "_" + $dat + ".log"
$FicRes     = $RepLog + $ScriptName + "_" + $dat + "_NTP.csv"
$LineSep    = "=" * 70

### Si le fichier LOG n'existe pas on le crée à vide
$Line = ">> DEBUT script récupération des statistiques de performances <<"
If (!(Test-Path $FicLog)) {
	$Line = (Get-Date -format G) + " " + $Line
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Line
} Else {
	LogTrace ($Line)
}

$tabDCexc = @() ; $tabDCinc = @() ; $tabClustExc = @() ; $tabClustInc = @() ; $tabESXexc = @() ; $tabESXinc = @() ; $tabTagexc = @() ; $tabTaginc = @()
$DCexc		= $DCexc.ToUpper().Trim()
$DCinc		= $DCinc.ToUpper().Trim()
$ClustExc	= $ClustExc.ToUpper().Trim()
$ClustInc	= $ClustInc.ToUpper().Trim()
$ESXexc		= $ESXexc.ToUpper().Trim()
$ESXinc		= $ESXinc.ToUpper().Trim()
$TagExc		= $TagExc.ToUpper().Trim()
$TagInc		= $TagInc.ToUpper().Trim()

If ($DCexc -eq "" -or $DCexc -eq "NONE") {  $DCexc = "AUCUN" }
If ($DCexc -ne "AUCUN") {  $tabDCexc = $DCexc.split(",") }
If ($DCinc -eq "" -or $DCinc -eq "NONE") {  $DCinc = "TOUS" }
If ($DCinc -ne "TOUS") {  $DCinc = $DCinc.split(",") }

If ($ClustExc -eq "" -or $ClustExc -eq "NONE") {  $ClustExc = "AUCUN" }
If ($ClustExc -ne "AUCUN") {  $tabClustExc = $ClustExc.split(",") }
If ($ClustInc -eq "" -or $ClustInc -eq "NONE") {  $ClustInc = "TOUS" }
If ($ClustInc -ne "TOUS") {  $tabClustInc = $ClustInc.split(",") }

If ($ESXexc -eq "" -or $ESXexc -eq "NONE") {  $ESXexc = "AUCUN" }
If ($ESXexc -ne "AUCUN") {  $tabESXexc = $ESXexc.split(",") }
If ($ESXinc -eq "" -or $ESXinc -eq "NONE") {  $ESXinc = "TOUS" }
If ($ESXinc -ne "TOUS") {  $tabESXinc = $ESXinc.split(",") }

If ($TagExc -eq "" -or $TagExc -eq "NONE") {  $TagExc = "AUCUN" }
If ($TagExc -ne "AUCUN") {  $tabTagexc = $TagExc.split(",") }
If ($TagInc -eq "" -or $TagInc -eq "NONE") {  $TagInc = "TOUS" }
If ($TagInc -ne "TOUS") {  $tabTaginc = $TagInc.split(",") }

LogTrace ("Fichier liste des vCenter à traiter .. : $vCenters")
LogTrace ("DATACENTER à exclure ..................: $DCexc")
LogTrace ("DATACENTER à prendre en compte ........: $DCinc")
LogTrace ("CLUSTER à exclure .................... : $ClustExc")
LogTrace ("CLUSTER à prendre en compte .......... : $ClustInc")
LogTrace ("ESX à exclure ........................ : $ESXexc")
LogTrace ("ESX à prendre en compte .............. : $ESXinc")
LogTrace ("Tags à exclure ........................: $TagExc")
LogTrace ("Tags à prendre en compte ..............: $TagInc")
LogTrace ("Action à réaliser .....................: $Action")
LogTrace ($LineSep + $vbcrlf)

$TabVcc   	= $vCenters.split(",")
$TabESXexc  = $ESXexc.split(",")
$TabTag		= $TagInc.split(",")

### Définition des entêtes du fichier de sortie
If ($Action -eq "SET") { $Entete  = "vCenter;Datacenter;Cluster;ESX;TAGS;NTP Server (AVANT);NTP Server (APRES);NTP Daemon (AVANT);NTP Daemon (APRES);vmKernel Gateway;Résultat" } Else { $Entete  = "vCenter;Datacenter;Cluster;ESX;TAGS;NTP Server;NTP Daemon;vmKernel Gateway;Résultat" }
Out-File -filepath $FicRes  -encoding UTF8 -inputobject $Entete

ForEach ($vCenter in $TabVcc) {
	LogTrace ("DEBUT du traitement du vCenter $vCenter")
	Write-Host "DEBUT du traitement du vCenter " -NoNewLine
	Write-Host "$vCenter... ".ToUpper() -ForegroundColor Yellow -NoNewLine
	Write-Host "En cours" -ForegroundColor Green	

	$rccnx = Connect-VIServer -Server $vcenter -Protocol https -Credential $Credential

	$topCnxVcc = "0"
	If ($rccnx -ne $null) {
		If ($rccnx.Isconnected) {
			$topCnxVcc = "1"
		}
	}

	If ($topCnxVcc -ne "1") {
		LogTrace ("ERREUR: Connexion KO au vCenter $vCenter => Arrêt du script")
		Write-Host "ERREUR: Connexion KO au vCenter $vCenter => Arrêt du script" -ForegroundColor White -BackgroundColor Red
		$rc += 1
		Exit $rc
	} Else {
		LogTrace ("SUCCES: Connexion OK au vCenter $vCenter" + $vbcrlf)
		Write-Host "SUCCES: Connexion OK au vCenter $vCenter`r`n" -ForegroundColor Black -BackgroundColor Green
	}

	$noDatacenter = 0
	$oDatacenters = Get-Datacenter | Sort Name
	$Datacenter_Counter = $oDatacenters.Count
	ForEach($DC in $oDatacenters){
		### Vérification de la liste noire
		If ($tabDCexc -like "*$DC*") {
			LogTrace ("Exclusion du DATACENTER $DC => ByPass du DATACENTER")
			Write-Host "Exclusion du DATACENTER $DC => ByPass du DATACENTER" -ForegroundColor White -BackgroundColor Red
		Continue }

		### Vérification de la liste blanche
		If ($tabDCinc.length -ne 0 -and $tabDCinc -notlike "*$DC*") {
			LogTrace ("DATACENTER $DC absent des Datacenters à prendre en compte => ByPass")
			Write-Host "DATACENTER $DC absent des Datacenters à prendre en compte => ByPass" -ForegroundColor White -BackgroundColor Red
		Continue }

		$noDatacenter += 1
		LogTrace ("Traitement du DATACENTER $DC n°$noDatacenter sur $Datacenter_Counter" + $vbcrlf)
		Write-Host "Traitement du DATACENTER [#$noDatacenter/$Datacenter_Counter] " -NoNewLine
		Write-Host "$DC... ".ToUpper() -ForegroundColor Yellow -NoNewLine
		Write-Host "En cours" -ForegroundColor Green	

		$noCluster = 0
		$oClusters = Get-Cluster -Location $DC | Sort Name
		$Cluster_Counter = $oClusters.Count

		ForEach($Cluster in Get-Cluster -Location $DC){
			### Récupération de la valeur du TAG par cluster
			$ClusterTAG = Get-TagAssignment -Entity $Cluster

			### Vérification de la liste noire CLUSTER
			If ($tabClustExc -like "*$Cluster*") {
				LogTrace ("Exclusion du CLUSTER $Cluster => ByPass du CLUSTER")
				Write-Host "Exclusion du CLUSTER $Cluster => ByPass du CLUSTER" -ForegroundColor White -BackgroundColor Red
			Continue }

			### Vérification de la liste blanche CLUSTER
			If ($tabClustInc.length -ne 0 -and $tabClustInc -notlike "*$Cluster*") {
				LogTrace ("CLUSTER $Cluster absent des clusters à prendre en compte => ByPass")
				Write-Host "CLUSTER $Cluster absent des clusters à prendre en compte => ByPass" -ForegroundColor White -BackgroundColor Red
			Continue }

			$no_VM = $no_ESX = $NULL
			$noCluster += 1

			LogTrace ("Traitement du CLUSTER $Cluster n°$noCluster sur $Cluster_Counter")
			Write-Host "Traitement du CLUSTER [#$noCluster/$Cluster_Counter] " -NoNewLine
			Write-Host "$Cluster... ".ToUpper() -ForegroundColor Yellow -NoNewLine
			Write-Host "En cours" -ForegroundColor Green	

			### Calcul du nombre d'ESX total
			$oESX = Get-VMHost -Location $Cluster | Sort Name
			$ESX_Counter = $oESX.Count

			ForEach($ESX in Get-VMHost -Location $Cluster) {
				If ($TagInc -eq "TOUS") { $oESXwithTAG = Get-TagAssignment -Entity $ESX | Where {$_.Tag.Name -Like "*"} } Else { $oESXwithTAG = Get-TagAssignment -Entity $ESX | Where {$_.Tag.Name -Like $TabTag} }
				$oESXwithTAG = Get-TagAssignment -Entity $ESX | Select -ExpandProperty Tag
				
				### Contrôle de la présence TAG + ESX connectés
				If (( $Action -eq "SET") -and ($oESXwithTAG -ne $NULL)) {
					### Vérification de la liste noire ESX
					If ($tabESXexc -like "*$ESX*") {
						LogTrace ("Exclusion de l'ESX '$ESX' => ByPass de l'hôte")
						Write-Host "Exclusion de l'ESX '$ESX' => ByPass de l'hôte" -ForegroundColor White -BackgroundColor Red
					Continue }

					### Vérification de la liste blanche ESX
					If ($tabESXinc.length -ne 0 -and $tabESXinc -notlike "*$ESX*") {
						LogTrace ("ESX '$ESX' absent des hôtes à prendre en compte => ByPass")
						Write-Host "ESX '$ESX' absent des hôtes à prendre en compte => ByPass" -ForegroundColor White -BackgroundColor Red
					Continue }

					### Récupération des valeurs NTP
					$Gateway = Get-VmHostNetwork -Host $ESX  -Server $vCenter | Select VMkernelGateway -ExpandProperty VirtualNic | Where {$_.VMotionEnabled} | Select -ExpandProperty VMkernelGateway
					$NTPServer_x = Get-VMHostNtpServer -VMHost $ESX -Server $vCenter
					If ((Get-VmHostService -VMHost $ESX  -Server $vCenter | Where-Object {$_.key -eq "ntpd"}).Running -eq "True") {
						$NTPd_x = "Running"	}
					Else {
						$NTPd_x = "Not Running"	}

					If ($NTPServer -ne $Gateway) {
						### Implémentation des valeurs NTP
						Get-VMHost | Get-VmHostService -VMHost $ESX -Refresh -Server $vCenter | Where-Object {$_.key -eq "ntpd"} | Stop-VMHostService –Confirm:$False
						Remove-VmHostNtpServer -NtpServer $NTPServer_x -VMHost $ESX -Server $vCenter –Confirm:$False
						Add-VmHostNtpServer -NtpServer $Gateway -VMHost $ESX -Server $vCenter –Confirm:$False
						Get-VMHost | Get-VMHostFirewallException -VMHost $ESX -Refresh -Server $vCenter | where {$_.Name -eq "NTP client"} | Set-VMHostFirewallException -Enabled:$True
						Get-VMhost | Get-VmHostService -VMHost $ESX -Refresh -Server $vCenter | Where-Object {$_.key -eq "ntpd"} | Set-VMHostService -policy "automatic" –Confirm:$False
						Get-VMHost | Get-VmHostService -VMHost $ESX -Refresh -Server $vCenter | Where-Object {$_.key -eq "ntpd"} | Start-VMHostService –Confirm:$False

						### Vérification des changements
						$NTPServer_y = Get-VMHostNtpServer -VMHost $ESX -Server $vCenter
						$NTPd_y = Get-VMHostService -VMHost $ESX -Refresh -Server $vCenter | where {$_.Key -eq 'ntpd'}
						If ($NTPServer_y -eq $Gateway -AND $NTPd_y.Running -eq $True) {	$Result = "OK" }
						Else { $Result = "NOK" }
					}
					$Result = "OK"
					
					LogTrace ("MISE A JOUR des informations dans le fichier de sortie pour l'ESX '$ESX'" + $vbcrlf)
					Write-Host "Mise à jour du fichier de sortie pour l'ESX '$ESX'" -NoNewLine -ForegroundColor White
					If ($Result -eq "OK") { Write-Host " $Result" -ForegroundColor Green } Else { Write-Host " $Result" -ForegroundColor Red }
					$RecordOut = "$vCenter;$Datacenter;$Cluster;$ESX;$oESXwithTAG;$NTPServer_x;$NTPServer_y;$NTPd_x;$NTPd_y;$Gateway;$Result"
					Out-File -filepath $FicRes -encoding UTF8 -inputobject $RecordOut -Append
				}
				Else {
					### Vérification de la liste noire ESX
					If ($tabESXexc -like "*$ESX*") {
						LogTrace ("Exclusion de l'ESX '$ESX' => ByPass de l'hôte")
						Write-Host "Exclusion de l'ESX '$ESX' => ByPass de l'hôte" -ForegroundColor White -BackgroundColor Red
					Continue }

					### Vérification de la liste blanche ESX
					If ($tabESXinc.length -ne 0 -and $tabESXinc -notlike "*$ESX*") {
						LogTrace ("ESX '$ESX' absent des hôtes à prendre en compte => ByPass")
						Write-Host "ESX '$ESX' absent des hôtes à prendre en compte => ByPass" -ForegroundColor White -BackgroundColor Red
					Continue }
					
					### Récupération des valeurs NTP
					$Gateway = Get-VmHostNetwork -Host $ESX  -Server $vCenter | Select VMkernelGateway -ExpandProperty VirtualNic | Where {$_.VMotionEnabled} | Select -ExpandProperty VMkernelGateway
					$NTPServer = Get-VMHostNtpServer -VMHost $ESX -Server $vCenter
					If ((Get-VmHostService -VMHost $ESX  -Server $vCenter | Where-Object {$_.key -eq "ntpd"}).Running -eq "True") {
						$NTPd = "Running"	}
					Else {
						$NTPd = "Not Running"	}
					
					If (($NTPServer -eq $Gateway) -AND ($NTPd -eq "Running")) { $Result = "OK" } Else { $Result = "NOK" }
					
					LogTrace ("MISE A JOUR des informations dans le fichier de sortie pour l'ESX '$ESX'" + $vbcrlf)
					Write-Host "Mise à jour du fichier de sortie pour l'ESX '$ESX'" -NoNewLine -ForegroundColor White
					If ($Result -eq "OK") { Write-Host " $Result" -ForegroundColor Green } Else { Write-Host " $Result" -ForegroundColor Red }
					$RecordOut = "$vCenter;$Datacenter;$Cluster;$ESX;$oESXwithTAG;$NTPServer;$NTPd;$Gateway;$Result"
					Out-File -filepath $FicRes -encoding UTF8 -inputobject $RecordOut -Append
				}
			}
		}
	}
	LogTrace ("DECONNEXION et FIN du traitement depuis le vCenter $vCenter`r`n")
	Disconnect-VIServer -Server $vCenter –Force –Confirm:$False
}