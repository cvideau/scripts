$global:MemoryInstalled			= $NULL
$global:MemoryConsumed			= $NULL
$global:ESXCoreMhz 				= $NULL
$global:CPUMHzAvg 				= $NULL
$global:MemoryAllocated 		= $NULL
$global:ClusterFreeDiskspaceTB	= $NULL

Function Get-FreeSpaceACP {
	Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$ClusterName, [Int]$oPctSecuDS, [Int]$oCapaMiniVM )
	ForEach ($oDS in Get-Cluster $Cluster | Get-Datastore | Where {$_.Accessible -eq "Mounted" -And $_.Name -NotLike "*datastore*" -And ($_.FreeSpaceMB - $_.CapacityMB * ($PctSecuDS/100)) -ge ($CapaMiniVM*1024)} ) {
		$ClusterFreeDiskspaceMB += ($oDS.FreeSpaceMB - $oDS.CapacityMB * ($PctSecuDS/100))
	}
	$global:ClusterFreeDiskspaceTB = [math]::round(($ClusterFreeDiskspaceMB/1024000), 1)
}

Function Get-VMHostInventory {
	Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	$hoststat = "" | Select HostName, MemoryInstalled, MemoryAllocated, MemoryConsumed, CPUMax, CPUAvg, CPUMin, CPUMHzMax, CPUMHzAvg, CPUMHzMin
	$hoststat.HostName = $ESX.name
	
	If ($HistoDays -eq "0") {
		$statcpu = Get-Stat -Entity $ESX -Realtime -stat cpu.usage.average
		$statcpuMHz = Get-Stat -Entity $ESX -Realtime -stat cpu.usagemhz.average
		$statmemconsumed = Get-Stat -Entity $ESX -Realtime -stat mem.consumed.average | Measure-Object -Property value -Average | Select-Object -ExpandProperty Average
	} Else {
		$statcpu = Get-Stat -Entity $ESX -Start $Start -Finish $Finish -MaxSamples 100 -stat cpu.usage.average
		$statcpuMHz = Get-Stat -Entity $ESX -Start $Start -Finish $Finish -MaxSamples 100 -stat cpu.usagemhz.average
		$statmemconsumed = Get-Stat -Entity $ESX -Start $Start -Finish $Finish -MaxSamples 100 -stat mem.consumed.average | Measure-Object -Property value -Average | Select-Object -ExpandProperty Average
	}

	Get-VMHost $ESX | Get-VM | ?{$_.PowerState -match 'PoweredOn'} | %{$statmemallocated=$statmemallocated+$_.MemoryGB}
	$statmeminstalled = Get-VMHost $ESX | Select MemoryTotalGB
	$statmeminstalled = $statmeminstalled.MemoryTotalGB

	$cpu = $statcpu | Measure-Object -Property value -Average -Maximum -Minimum
	$cpuMHz = $statcpuMHz | Measure-Object -Property value -Average -Maximum -Minimum

	[int]$CPUAvg = "{0:N0}" -f ($cpu.Average)
	$global:CPUMHzAvg = [math]::round(($cpuMHz.Average/1024), 0)
	$global:MemoryAllocated = $statmemallocated
	$global:MemoryConsumed = [math]::round(($statmemconsumed/1024000), 2)
	$global:MemoryInstalled = [math]::round($statmeminstalled, 0)

	#CPU info
	$ESXCpuSockets = $ESX.ExtensionData.Summary.Hardware.NumCpuPkgs * ($ESX.ExtensionData.Summary.Hardware.NumCpuCores/$ESX.ExtensionData.Summary.Hardware.NumCpuPkgs)
	$global:ESXCoreMhz = [math]::round(($ESX.CPUTotalMhz/1024), 0)
}

Add-PSsnapin VMware.VimAutomation.Core

$HistoDays			= $args[0] # Nombre de jours glissants pour la récupération des stats
$vCenters			= $args[1] # Nom du vCenter
$DCexc				= $args[2] # Datacenter exclus
$DCinc				= $args[3] # Datacenter inclus
$ClustExc			= $args[4] # Clusterlusters exclus
$ClustInc			= $args[5] # CPUlusters inclus
$ESXexc				= $args[6] # ESX exclus
$ESXinc				= $args[7] # ESX inclus
$TagExc				= $args[8] # TAGs exclus
$TagInc				= $args[9] # TAGs inclus
[Int]$PctSecuDS		= $args[10] # Pourcentage du seuil de sécurité de chacun des Datastores
[Int]$CapaMiniVM	= $args[11] # Capacité disque minimale pour la création d'une VM

# Bouchon pour tests
#$HistoDays	= 1
#$vCenters	= "swmuzv1vcszc.zres.ztech"
#$ClustExc	= "AUCUN"
#$DCexc		= "AUCUN"
#$DCinc		= "TOUS"
#$ESXexc	= "AUCUN"
#$ClustInc	= "CL_MU_GSA_Z2"
#$ESXinc	= "TOUS"
#$TagInc	= "Z_MU_ESX_Production"
#$TagExc	= "AUCUN"
#$PctSecuDS	= 80
#$CapaMiniVM= 60

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
$FicRes     = $RepLog + $ScriptName + "_" + $dat + "_ClusterLoad.csv"
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

LogTrace ("Historique des statistiques en jours.. : $HistoDays")
LogTrace ("Fichier liste des vCenter à traiter .. : $vCenters")
LogTrace ("DATACENTER à exclure ..................: $DCexc")
LogTrace ("DATACENTER à prendre en compte ........: $DCinc")
LogTrace ("CLUSTER à exclure .................... : $ClustExc")
LogTrace ("CLUSTER à prendre en compte .......... : $ClustInc")
LogTrace ("ESX à exclure ........................ : $ESXexc")
LogTrace ("ESX à prendre en compte .............. : $ESXinc")
LogTrace ("Tags à exclure ........................: $TagExc")
LogTrace ("Tags à prendre en compte ..............: $TagInc")
LogTrace ("Valeur de sécurité Datastore ..........: $PctSecuDS %")
LogTrace ("Capacité stockage mini VM .............: $CapaMiniVM Mo")
LogTrace ($LineSep + $vbcrlf)

$TabVcc   	= $vCenters.split(",")
$TabESXexc  = $ESXexc.split(",")
$TabTag		= $TagInc.split(",")

### Définition des entêtes du fichier de sortie
$Entete  = "vCenter;Datacenter;Cluster;TAG Cluster;NB VM;NB ESX PROD;NB ESX MAINT;NB ESX CONSTR;NB VM/ESX;CAPA RAM ESX PROD (To);ALLOC. RAM PROD (%);ALLOC. RAM PREV. (%);UTIL. RAM PROD (%);UTIL. RAM PREV. (%);CAPA CPU ESX PROD (GHz);UTIL. CPU PROD(%);UTIL. CPU PREV.(%);CAPA DISK ESX PROD (To);ALLOC. DISK PROD(%);CAPA DISK PROD ACP (To)"
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

	$Start = (Get-Date).AddDays(-$HistoDays)
	$Finish = Get-Date

	$noDatacenter = 0
	$oDatacenters = Get-Datacenter | Sort Name
	$Datacenter_Counter = $oDatacenters.Count
	ForEach($DC in $oDatacenters){
		### Vérification de la liste noire
		If ($tabDCexc -contains $DC) {
			LogTrace ("Exclusion du DATACENTER $DC => ByPass du DATACENTER")
			Write-Host "Exclusion du DATACENTER $DC => ByPass du DATACENTER" -ForegroundColor White -BackgroundColor Red
		Continue }

		### Vérification de la liste blanche
		If ($tabDCinc.length -ne 0 -and $tabDCinc -notcontains $DC) {
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
			$StartTime = (Get-Date)
			### Récupération de la valeur du TAG par cluster
			$ClusterTAG = Get-TagAssignment -Entity $Cluster
			
			### Vérification de la liste noire CLUSTER
			If ($tabClustExc -contains $Cluster) {
				LogTrace ("Exclusion du CLUSTER $Cluster => ByPass du CLUSTER")
				Write-Host "Exclusion du CLUSTER $Cluster => ByPass du CLUSTER" -ForegroundColor White -BackgroundColor Red
			Continue }

			### Vérification de la liste blanche CLUSTER
			If ($tabClustInc.length -ne 0 -and $tabClustInc -notcontains $Cluster) {
				LogTrace ("CLUSTER $Cluster absent des clusters à prendre en compte => ByPass")
				Write-Host "CLUSTER $Cluster absent des clusters à prendre en compte => ByPass" -ForegroundColor White -BackgroundColor Red
			Continue }

			$no_VM = $no_ESX = $NULL
			$noCluster += 1

			LogTrace ("Traitement du CLUSTER $Cluster n°$noCluster sur $Cluster_Counter")
			Write-Host "Traitement du CLUSTER [#$noCluster/$Cluster_Counter] " -NoNewLine
			Write-Host "$Cluster... ".ToUpper() -ForegroundColor Yellow -NoNewLine
			Write-Host "En cours" -ForegroundColor Green	
			
			### Récupération des métriques au niveau CLUSTER
			$ClusterCPUCores 			= $Cluster.ExtensionData.Summary.NumCpuCores
			$ClusterCapacityDiskspaceGB	= [math]::round(($Cluster | Get-Datastore | Where-Object {$_.Extensiondata.Summary.MultipleHostAccess -eq $True} | Measure-Object -Property CapacityGB -Sum).Sum/1024, 0)
			$ClusterFreeDiskspaceGB		= [math]::round(($Cluster | Get-Datastore | Where-Object {$_.Extensiondata.Summary.MultipleHostAccess -eq $True} | Measure-Object -Property FreeSpaceGB -Sum).Sum/1024, 0)
			$ClusterUsedDiskspaceGB		= [math]::round(($ClusterCapacityDiskspaceGB - $ClusterFreeDiskspaceGB), 0)
			Get-FreeSpaceACP -ClusterName $Cluster -oPctSecuDS $PctSecuDS -oCapaMiniVM $CapaMiniVM
			
			### Calcul du nombre d'ESX total
			$oESX = Get-VMHost -Location $Cluster | Sort Name
			$ESX_Counter = $oESX.Count
			
			### Calcul du nombre d'ESX en mode MAINTENANCE
			$oESXMaintMode = Get-VMHost -State "Maintenance" -Location $Cluster
			$ESXMaintMode_Counter = $oESXMaintMode.Count

			### Récupération des statistiques de performances des ESX en PRODUCTION ET CONNECTED
			$ESXCoreMhz = $MemoryConsumed = $MemoryInstalled = $CPUAvg = $CPUMHzAvg = $MemoryAllocated = $NULL
			$RAMCapaPROD = $RAMUsagePROD = $RAMAllocPROD = $CPUCapaPROD = $CPUUsagePROD = $NULL
			$ESXwithTag_Counter = $ESXwithoutTag_Counter = $NULL

			ForEach($ESX in Get-VMHost -Location $Cluster) {
				If ($TagInc -eq "TOUS") { $oESXwithTAG = Get-TagAssignment -Entity $ESX | Where {$_.Tag.Name -Like "*"} } Else { $oESXwithTAG = Get-TagAssignment -Entity $ESX | Where {$_.Tag.Name -Like $TabTag} }
				If ($oESXwithTAG -ne $NULL) { $ESXwithTag_Counter += 1 } Else { $ESXwithoutTag_Counter += 1	}
					
				# Contrôle de la présence TAG + ESX connectés
				If (( $ESX.ConnectionState -Match "Connected") -and ($oESXwithTAG -ne $null)) {
					### Vérification de la liste noire ESX
					If ($tabESXexc -contains $ESX) {
						LogTrace ("Exclusion de l'ESX $ESX => ByPass de l'hôte")
						Write-Host "Exclusion de l'ESX $ESX => ByPass de l'hôte" -ForegroundColor White -BackgroundColor Red
					Continue }

					### Vérification de la liste blanche ESX
					If ($tabESXinc.length -ne 0 -and $tabESXinc -notcontains $ESX) {
						LogTrace ("ESX $ESX absent des hôtes à prendre en compte => ByPass")
						Write-Host "ESX $ESX absent des hôtes à prendre en compte => ByPass" -ForegroundColor White -BackgroundColor Red
					Continue }

					$no_ESX += 1
					$no_VM += (Get-VM -Location $ESX | ?{$_.PowerState -match 'PoweredOn'}).count

					LogTrace ("Récupération des statistiques de performances de l'ESX (TAG '$oESXwithTAG' PRODUCTION) $ESX n°$no_ESX sur $ESXwithTag_Counter du CLUSTER $Cluster")
					Write-Host "Récup. des perf. de l'ESX " -NoNewLine
					Write-Host "[Prod]" -ForegroundColor Black -BackgroundColor White -NoNewLine
					Write-Host " [#$no_ESX/$ESX_Counter] " -NoNewLine
					Write-Host "$ESX...".ToUpper() -ForegroundColor Yellow -NoNewLine

					### Récupération des statistiques actuelles (Par ESX)
					Get-VMHostInventory -vmHost $ESX
					$RAMCapaPROD 	+= $MemoryInstalled	# Total de la capacité RAM des ESX en PRODUCTION TAGués et CONNECTES (To)
					$RAMUsagePROD 	+= $MemoryConsumed	# Total de la RAM utilisée des ESX en PRODUCTION TAGués et CONNECTES (To)
					$RAMAllocPROD	+= $MemoryAllocated	# Total de la RAM allouée aux VM sur les ESX en PRODUCTION TAGués et CONNECTES (Go)
					$CPUCapaPROD 	+= $ESXCoreMhz		# Total de la capacité CPU des ESX en PRODUCTION TAGués et CONNECTES (GHz)
					$CPUUsagePROD 	+= $CPUMHzAvg		# Total de la CPU utilisée des ESX en PRODUCTION TAGués et CONNECTES (GHz)
					Write-Host " OK" -ForegroundColor Green
				}
			}
			$RAMTauxAllocPROD 	= [math]::round(($RAMAllocPROD/$RAMCapaPROD)*100, 0)	# Total de la RAM allouée aux VM démarrées par ESX TAGués et CONNECTES(To) / Total de la capacité RAM des ESX en PRODUCTION TAGués  et CONNECTES (To) - Exprimé en %
			$RAMTauxUsagePROD 	= [math]::round(($RAMUsagePROD/$RAMCapaPROD)*100, 0)	# Total de la RAM utilisée des ESX en PRODUCTION TAGués et CONNECTES(To) / Total de la capacité RAM des ESX en PRODUCTION TAGués  et CONNECTES(To) - Exprimé en %
			$CPUTauxUsagePROD	= [math]::round(($CPUUsagePROD/$CPUCapaPROD)*100, 0)	# Total de la CPU utilisée des ESX en PRODUCTION TAGués et CONNECTES (GHz) / Total de la capacité CPU des ESX en PRODUCTION TAGués et CONNECTES (GHz) - Exprimé en %
			$DISKCapaPROD 		= $ClusterCapacityDiskspaceGB							# Total de la capacité DISK du cluster (To)
			$DISKUsagePROD 		= $ClusterUsedDiskspaceGB								# Total de la capacité DISK utilisée du cluster (To)
			$DISKTauxUsagePROD 	= [math]::round(($DISKUsagePROD/$DISKCapaPROD)*100, 0)	# Total de la capacité DISK utilisée du cluster (To) / Total de la capacité DISK du cluster (To) - Exprimé en %


			### Récupération des statistiques de performances des ESX en CONSTRUCTION ou en MAINTENANCE
			$ESXCoreMhz = $MemoryInstalled = $MemoryAllocated = $NULL
			$RAMCapaPREV = $RAMAllocPREV = $CPUCapaPREV = $NULL
				
			ForEach($ESX in Get-VMHost -Location $Cluster) {
				If ($TagInc -eq "TOUS") { $oESXwithTAG = Get-TagAssignment -Entity $ESX | Where {$_.Tag.Name -Like "*"} } Else { $oESXwithTAG = Get-TagAssignment -Entity $ESX | Where {$_.Tag.Name -Like $TabTag} }
				If (!(($ESX.ConnectionState -Match "Connected") -and ($oESXwithTAG -ne $null))) {
					### Vérification de la liste noire ESX
					If ($tabESXexc -contains $ESX) {
						LogTrace ("Exclusion de l'ESX $ESX => ByPass de l'hôte")
						Write-Host "Exclusion de l'ESX $ESX => ByPass de l'hôte" -ForegroundColor White -BackgroundColor Red
					Continue }

					### Vérification de la liste blanche ESX
					If ($tabESXinc.length -ne 0 -and $tabESXinc -notcontains $ESX) {
						LogTrace ("ESX $ESX absent des hôtes à prendre en compte => ByPass")
						Write-Host "ESX $ESX absent des hôtes à prendre en compte => ByPass" -ForegroundColor White -BackgroundColor Red
					Continue }

					$no_ESX += 1
					$no_VM += (Get-VM -Location $ESX | ?{$_.PowerState -match 'PoweredOn'}).count

					LogTrace ("Récupération des statistiques de performances de l'ESX (en Construction ou en Maintenance) $ESX n°$no_ESX sur $ESXwithTag_Counter du CLUSTER $Cluster")
					Write-Host "Récup. des perf. de l'ESX " -NoNewLine
					Write-Host "[Hors Prod]" -ForegroundColor Black -BackgroundColor White -NoNewLine
					Write-Host " [#$no_ESX/$ESX_Counter] " -NoNewLine
					Write-Host "$ESX...".ToUpper() -ForegroundColor Yellow -NoNewLine

					### Récupération des statistiques prévisionnelles (Par ESX)
					Get-VMHostInventory -vmHost $ESX
					$RAMCapaPREV 	+= $MemoryInstalled	# Total de la capacité RAM des ESX en non PRODUCTION ou non CONNECTES (To)
					$CPUCapaPREV 	+= $ESXCoreMhz		# Total de la capacité CPU des ESX en non PRODUCTION ou non CONNECTES (GHz)
					Write-Host " OK" -ForegroundColor Green
				}
			}
				
			$RAMTauxAllocPREV 	= [math]::round(($RAMAllocPROD/($RAMCapaPROD+$RAMCapaPREV))*100, 0)	# Total de la RAM allouée aux VM démarrées par ESX TAGués et CONNECTES(To) / Total de la capacité RAM de TOUS les ESX du cluster (To) - Exprimé en %
			$RAMTauxUsagePREV 	= [math]::round(($RAMUsagePROD/($RAMCapaPROD+$RAMCapaPREV))*100, 0)	# Total de la RAM utilisée des ESX en PRODUCTION TAGués et CONNECTES(To) / Total de la capacité RAM de TOUS les ESX du cluster (To) - Exprimé en %
			$CPUTauxUsagePREV 	= [math]::round(($CPUUsagePROD/($CPUCapaPROD+$CPUCapaPREV))*100, 0)	# Total de la CPU utilisée des ESX en PRODUCTION TAGués et CONNECTES (GHz) / Total de la capacité CPU de TOUS les ESX du cluster (GHz) - Exprimé en %
			
			LogTrace ("MISE A JOUR des informations dans le fichier de sortie pour le CLUSTER $Cluster" + $vbcrlf)
			Write-Host "Mise à jour du fichier de sortie pour le CLUSTER " -NoNewLine
			Write-Host "$Cluster... "  -ForegroundColor Yellow -NoNewLine
			$RecordOut  = "$vCenter;$DC;$Cluster;$ClusterTAG;$no_VM;$ESXwithTag_Counter;$ESXMaintMode_Counter;$ESXwithoutTag_Counter;" + [math]::round($no_VM/($ESXwithTag_Counter - $ESXMaintMode_Counter), 0) + ";" + [math]::round($RAMCapaPROD/1024, 1) + ";$RAMTauxAllocPROD;$RAMTauxAllocPREV;$RAMTauxUsagePROD;$RAMTauxUsagePREV;$CPUCapaPROD;$CPUTauxUsagePROD;$CPUTauxUsagePREV;$DISKCapaPROD;$DISKTauxUsagePROD;$ClusterFreeDiskspaceTB"
			Out-File -Filepath $FicRes  -Encoding UTF8 -InputObject $RecordOut -Append
			$EndTime = (Get-Date)
			[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
			[String]$ElapsedTimeESX = [math]::round($ElapsedTime/$no_ESX, 0)
			Write-Host "OK" -ForegroundColor Green
			Write-Host "Temps d'exécution " -NoNewLine
			Write-Host "${ElapsedTime} " -ForegroundColor White -NoNewLine
			Write-Host "secondes ($ElapsedTimeESX secondes par ESX)`r`n"
		}
	}
	LogTrace ("DECONNEXION et FIN du traitement depuis le vCenter $vCenter`r`n")
	Disconnect-VIServer -Server $vCenter –Force –Confirm:$False
}