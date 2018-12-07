$global:MemoryInstalled			= $NULL
$global:MemoryConsumed			= $NULL

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
$FicRes     = $PathScript + "Overcommitting.xlsm"
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
LogTrace ($LineSep + $vbcrlf)

$TabVcc   	= $vCenters.split(",")
$TabESXexc  = $ESXexc.split(",")
$TabTag		= $TagInc.split(",")

### Définition des entêtes du fichier de sortie
$Excel = New-Object -ComObject Excel.Application
$ExcelWorkBook  = $Excel.WorkBooks.Open($FicRes)
$ExcelWorkSheet = $Excel.WorkSheets.item(1)
$ExcelWorkSheet.Activate()
$Excel.Visible = $False


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
			
			### Calcul du nombre d'ESX total
			$oESX = Get-VMHost -Location $Cluster | Sort Name
			$ESX_Counter = $oESX.Count
			$noCluster += 1

			LogTrace ("Traitement du CLUSTER $Cluster n°$noCluster sur $Cluster_Counter")
			Write-Host "Traitement du CLUSTER [#$noCluster/$Cluster_Counter] " -NoNewLine
			Write-Host "$Cluster... ".ToUpper() -ForegroundColor Yellow -NoNewLine
			Write-Host "En cours" -ForegroundColor Green	
			
			### Récupération des statistiques de performances des ESX en PRODUCTION ET CONNECTED
			$MemoryConsumed = $MemoryInstalled = $NULL

			ForEach($ESX in Get-VMHost -Location $Cluster) {
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
				
				$no_VM = (Get-VM -Location $ESX | ?{$_.PowerState -match 'PoweredOn'}).count
				$MemoryAllocated = 0
				$no_ESX += 1

				LogTrace ("Récupération des statistiques de performances de l'ESX (TAG '$oESXwithTAG' PRODUCTION) $ESX n°$no_ESX sur $ESXwithTag_Counter du CLUSTER $Cluster")
				Write-Host "Récup. des perf. de l'ESX " -NoNewLine
				Write-Host "$ESX...".ToUpper() -ForegroundColor Yellow -NoNewLine

				### Récupération des statistiques (Par ESX)
				If ($HistoDays -eq "0") {
					$statmemconsumed = Get-Stat -Entity $ESX -Realtime -stat mem.consumed.average | Measure-Object -Property value -Average | Select-Object -ExpandProperty Average
				} Else {
					$statmemconsumed = Get-Stat -Entity $ESX -Start $Start -Finish $Finish -MaxSamples 100 -stat mem.consumed.average | Measure-Object -Property value -Average | Select-Object -ExpandProperty Average
				}
				$statmeminstalled = Get-VMHost $ESX | Select MemoryTotalMB
				$statmeminstalled = $statmeminstalled.MemoryTotalMB
				$MemoryConsumed = [math]::round(($statmemconsumed/1024), 0)
				$MemoryInstalled = [math]::round($statmeminstalled, 0)
				Get-VMHost $ESX | Get-VM | %{$MemoryAllocated = $MemoryAllocated + $_.MemoryMB}

				$oESXwithTAG	= ($ESX | Get-TagAssignment).Tag.name
				$oESXState		= $ESX.ConnectionState
				Write-Host " OK" -ForegroundColor Green
				
				LogTrace ("MISE A JOUR des informations dans le fichier de sortie pour l'ESX $ESX" + $vbcrlf)
				Write-Host "Mise à jour du fichier de sortie pour l'ESX " -NoNewLine
				Write-Host "$ESX... "  -ForegroundColor Yellow -NoNewLine
				
				$ExcelWorkSheet.Cells.Item($no_ESX + 1, 1) = "$ESX" # Ligne, Colonne
				$ExcelWorkSheet.Cells.Item($no_ESX + 1, 2) = "$MemoryInstalled" # Ligne, Colonne
				$ExcelWorkSheet.Cells.Item($no_ESX + 1, 3) = "$MemoryConsumed" # Ligne, Colonne
				$ExcelWorkSheet.Cells.Item($no_ESX + 1, 4) = "$MemoryAllocated" # Ligne, Colonne
				$ExcelWorkSheet.Cells.Item($no_ESX + 1, 7) = "$Cluster" # Ligne, Colonne
				$ExcelWorkSheet.Cells.Item($no_ESX + 1, 8) = "$vCenter" # Ligne, Colonne
				$ExcelWorkSheet.Cells.Item($no_ESX + 1, 9) = "$no_VM" # Ligne, Colonne
				$ExcelWorkSheet.Cells.Item($no_ESX + 1, 10) = "$oESXwithTAG" # Ligne, Colonne
				$ExcelWorkSheet.Cells.Item($no_ESX + 1, 11) = "$oESXState" # Ligne, Colonne
				
				Write-Host "OK" -ForegroundColor Green
			}
		}
	}
	LogTrace ("DECONNEXION et FIN du traitement depuis le vCenter $vCenter`r`n")
	Disconnect-VIServer -Server $vCenter –Force –Confirm:$False
}
$ExcelWorkBook.Save()
$ExcelWorkBook.Close()
$Excel.Quit()