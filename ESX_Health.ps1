﻿<# Corrections apportées en 0.6
 * Ajout de l'export au format CSV
 * Création d'un fichier .lock lors de la première exécution du script qui interdit une nouvelle exécution
 * Ajout d'un mode DEBUG qui ajoute le n° de ligne du code dans le fichier LOG (Variable "$Mode_Debug". 0=Désactivé 1=Activé)
 * Ajout d'un bout de code de type "bouchon"
 * Ajout des nouveaux VCenters dans la commande SWITCH
 * Initialisation des variables _REF à null en début de chaque fonction
 * Correction de bugs mineurs
#>


### V0.6 (décembre 2018) - Développeur: Christophe VIDEAU
# Lien vers couleur Excel https://docs.microsoft.com/en-us/office/vba/images/colorin_za06050819.gif
# Lien vers couleur Write-Host https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-host?view=powershell-6
$WarningPreference = "SilentlyContinue"
Import-Module "VMware.VimAutomation.Core"

Import-Module VMware.VimAutomation.Core -WarningAction SilentlyContinue
Import-Module VMware.VimAutomation.Vds -WarningAction SilentlyContinue
Import-Module VMware.VimAutomation.License -WarningAction SilentlyContinue
Import-Module VMware.VimAutomation.Storage -WarningAction SilentlyContinue
Import-Module VMware.VimAutomation.HA -WarningAction SilentlyContinue

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
Add-PsSnapin VMware.VumAutomation
Add-PSSnapin -PassThru VMware.VimAutomation.Core
#$ErrorActionPreference = "SilentlyContinue"
Clear-Host

Write-Host "Développement (2018) Christophe VIDEAU - Version 0.6`r`n" -ForegroundColor White

### Déclaration des variables d'environnement
$VBCrLF							= "`r`n"
$L								= 0
$Format_DATE					= Get-Date -UFormat "%Y_%m_%d_%H_%M_%S"
$Mode_Var						= $Null
$Mode_Debug						= 1

### Déclaration des variables d'usage
$Global:ElapsedTime_Start		= 0
$Global:ElapsedTime_End			= 0
$Global:ESX_Compliant 			= 0
$Global:ESX_NotCompliant 		= 0
$Global:ESX_NotCompliant_Item	= 0
$Global:ESX_NotCompliant_Total	= 0
$Global:no_inc_ESX 				= 0
$Global:ESX_Counter 			= 0
$Global:Cluster_Counter 		= 0
$Global:Cluster_ID 				= 0
$Global:vCenter					= $Null
$Global:Cluster 				= $Null
$Global:ExcelLine_Start 		= 2			### Démarrage à la 2ème ligne du fichier Excel
$Global:ESX_Item_CheckType		= 21		### Nombre de paramètres (colonnes) vérifiés par ESXi
$Global:Cluster_Item_CheckType	= 11		### Nombre de paramètres (colonnes) vérifiés par clusters

### Déclaration des variables de fichiers
$Global:ScriptName				= [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition)
$Global:PathScript 				= ($myinvocation.MyCommand.Definition | Split-Path -parent) + "\"
$Global:RepLog     				= $PathScript + "LOG\"
$Global:File_TEMPLATE			= "ESX_Health_Modele.xlsx"
$Global:File_CSV 				= $RepLog + $ScriptName + "_REF.csv"
$Global:File_XLSX_DISCOVER_bckp	= "_REF_(BCKP_"+ $Format_DATE + ").xlsx"
$Global:File_LOG_DISCOVER_bckp	= "_REF_(BCKP_"+ $Format_DATE + ").log"
$Global:File_CSV_DISCOVER_bckp	= "_REF_(BCKP_"+ $Format_DATE + ").csv"
$Global:File_XLSX_DISCOVER_work	= "_REF-work.xlsx"
$Global:File_XLSX_DISCOVER		= "_REF.xlsx"
$Global:File_XLSX_ONESHOT		= "_ONE_SHOT.xlsx"
$Global:File_PROCESS_LOCK		= ".lock"
$Global:File_LOG_DISCOVER		= "_REF_work.log"
$Global:File_CSV_DISCOVER		= "_REF.csv"
$Global:File_LOG_ONESHOT		= "_ONE_SHOT.log"

### Déclaration des variabes de sécurité
$USER       	= "CSP_SCRIPT_ADM"
$KeyFile     	= "D:\Scripts\Credentials\key.crd"
$CredentialFile = "D:\Scripts\Credentials\vmware_adm.crd"
$key        	= Get-content $KeyFile
$PASSWORD       = Get-Content $CredentialFile | ConvertTo-SecureString -Key $key
$Credential 	= New-Object System.Management.Automation.PSCredential $USER, $PASSWORD

### Définition des variables selon les entêtes Excel
$Excel_Nom_ESX					= 1
$Excel_ENVironnement			= 2
$Excel_HARD_Constructeur		= 3
$Excel_HARD_Modele				= 4
$Excel_HARD_Num_Serie			= 5
$Excel_HARD_Version_BIOS		= 6
$Excel_HARD_Date_BIOS			= 7
$Excel_HARD_Carte_SD			= 8
$Excel_CONF_Datacenter			= 9
$Excel_CONF_Cluster				= 10
$Excel_CONF_Etat_ESX			= 11
$Excel_CONF_vCenter				= 12
$Excel_CONF_Balise				= 13
$Excel_CONF_Statut_ESX			= 14
$Excel_CONF_Version				= 15
$Excel_CONF_Date_Installation	= 16
$Excel_CONF_Serveurs_NTP		= 17
$Excel_CONF_Zone_Temps			= 18
$Excel_CONF_Heure_Locale		= 19
$Excel_CONF_Niveau_Compliance	= 20
$Excel_CONF_Etat_Demon_NTP		= 21
$Excel_SAN_Total_Chemins		= 22
$Excel_SAN_Total_LUNs			= 23
$Excel_SAN_Total_Chemins_MORTS	= 24
$Excel_SAN_Total_Chemins_ACTIFS	= 25
$Excel_SAN_Redondance_Chemins	= 26
$Excel_HOST_Current_CPU_Policy	= 27
$Excel_HOST_Mode_HA				= 28
$Excel_HOST_Hyperthreading		= 29
$Excel_HOST_EVC					= 30
$Excel_HOST_Etat_Alarmes		= 31
$Excel_HOST_TPS_Salting			= 32
$Excel_HOST_Larges_Pages_RAM	= 33
$Excel_NETWORK_VLAN				= 34
$Excel_NETWORK_Adaptateurs		= 35
$Excel_NETWORK_vMotion			= 36
$Excel_PSI_Eligibilite			= 37
$Excel_ILO_Version				= 38
$Excel_Conformite_Globale		= 39
$Excel_Conformite_Details		= 40
$Excel_TimeStamp				= 41

$Excel_Ref_Cluster				= 1
$Excel_Ref_ArrayVersion			= 2
$Excel_Ref_ENVironnement		= 3
$Excel_Ref_ArrayTimeZone		= 4
$Excel_Ref_ArrayBaseline		= 5
$Excel_Ref_ArrayNTPd			= 6
$Excel_Ref_ArraySAN				= 7
$Excel_Ref_ArrayLUN				= 8
$Excel_Ref_ArrayCPUPolicy		= 9
$Excel_Ref_ArrayHA				= 10
$Excel_Ref_ArrayHyperTh			= 11
$Excel_Ref_ArrayEVC				= 12
$Excel_Ref_ArrayAlarme			= 13
$Excel_Ref_ArrayTPS				= 14
$Excel_Ref_ArrayLPages			= 15
$Excel_Ref_ArrayVLAN			= 16
$Excel_Ref_ArrayLAN				= 17
$Excel_Ref_Conformite_Globale	= 18
$Excel_Ref_Conformite_Details	= 19
$Excel_Ref_TimeStamp			= 20


### Définition des variables d'utilisation Excel
$Excel_Couleur_Error			= 3
$Excel_Couleur_Background		= 15


### Fonction chargée de mesurer le temps de traitement
Function Get-ElapsedTime {
	$ElapsedTime_End = (Get-Date)
	$ElapsedTime = ($ElapsedTime_End - $ElapsedTime_Start).TotalSeconds
	$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
	$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
	If ($Sec.ToString().Length -eq 1) { $Sec = "0" + $Sec }
	Write-Host "[$($Min)min. $($Sec)sec]" -ForegroundColor White
}


### Fonction chargée de récupérer les valeurs du MATERIELS
Function Get-ESX_HARD { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "HARDWARE`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs HARDWARE")
	$ElapsedTime_Start = (Get-Date)
	
	$ESXCLI = Get-ESXCLI -VMHost $ESX.Name -Server $vCenter
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_Nom_ESX)	= $ESX.Name

	Switch -Wildcard ($vCenter)	{
		"*VCSZ*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "PRODUCTION" }
		"*VCSY*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "NON PRODUCTION" }
		"*VCSQ*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "CLOUD" }
		"*VCSZY*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement) 	= "SITES ADMINISTRATIFS" }
		"*VCSSA*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "BAC A SABLE" }
		"*VCS00*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "PVM" }
		Default	{ $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)		= "NA" }
	}
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement).Interior.ColorIndex = $Excel_Couleur_Background
	
	
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Constructeur)	= $ESXCLI.Hardware.Platform.Get.Invoke().VendorName
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Modele)			= $ESXCLI.Hardware.Platform.Get.Invoke().ProductName
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Num_Serie)		= $ESXCLI.Hardware.Platform.Get.Invoke().SerialNumber
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Version_BIOS)	= $ESX.ExtensionData.Hardware.BiosInfo.BiosVersion
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Date_BIOS)		= $ESX.ExtensionData.Hardware.BiosInfo.ReleaseDate
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Carte_SD)		= $ESXCLI.Storage.Core.Path.List.Invoke().AdapterTransportDetails | Where { $_.Device -eq "mpx.vmhba32:C0:T0:L0" }
	
	### Colorisation des cellules
	$ESXCLI = Get-ESXCLI -VMHost $ESX.Name -Server $vCenter
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_Nom_ESX).Interior.ColorIndex 			= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Constructeur).Interior.ColorIndex	= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Modele).Interior.ColorIndex 		= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Num_Serie).Interior.ColorIndex 		= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Version_BIOS).Interior.ColorIndex 	= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Date_BIOS).Interior.ColorIndex 		= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HARD_Carte_SD).Interior.ColorIndex 		= $Excel_Couleur_Background
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de CONFIGURATION
Function Get-ESX_CONFIG { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	
	### Initialisation des variables
	$Global:ESX_State 		= $Null
	$Global:oESXTAG 		= $Null
	$Global:oESX_NTP		= $Null
	$Global:InstallDate		= $Null
	$Global:UTC				= $Null
	$Global:UTC_Var1		= $Null
	$Global:UTC_Var2		= $Null
	$Global:Gateway			= $Null
	$Global:ComplianceLevel = $Null
	$Global:NTPD_status		= $Null

	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "CONFIGURATION`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs CONFIGURATION")
	$ElapsedTime_Start = (Get-Date)
	
	$ESXCLI = Get-ESXCLI -VMHost $ESX.Name -Server $vCenter
	$Global:oESXTAG = Get-VMHost -Name $ESX | Get-TagAssignment | Select Tag
	$Global:oESX_NTP = Get-VMHostNtpServer -VMHost $ESX.Name -Server $vCenter
	
	If ((Get-VmHostService -VMHost $ESX -Server $vCenter | Where-Object {$_.key -eq "ntpd"}).Running -eq "True") { $ntpd_status = "Running" } Else { $ntpd_status = "Not Running" }
	If ($ESXCLI.System.MaintenanceMode.Get.Invoke() -eq "Enabled") { $ESX_State = "Maintenance" } Else { $ESX_State = "Connected" }
	If ((Get-Compliance -Entity $ESX -Detailed -WarningAction "SilentlyContinue" | Where {$_.NotCompliantPatches -ne $Null} | Select Status).Count -gt 0) { $ComplianceLevel = "Baseline - Not compliant" } Else { $ComplianceLevel = "Baseline - Compliant" }
	$Global:UTC = Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}
	$Global:UTC_Var1 = $UTC.Config.DateTimeInfo.TimeZone.Name
	$Global:UTC_Var2 = $UTC.Config.DateTimeInfo.TimeZone.GmtOffset
	$Global:Gateway = Get-VmHostNetwork -Host $ESX -Server $vCenter | Select VMkernelGateway -ExpandProperty VirtualNic | Where { $_.VMotionEnabled } | Select -ExpandProperty VMkernelGateway -WarningAction "SilentlyContinue"
	
	New-VIProperty -Name EsxInstallDate -ObjectType VMHost -Value { Param($ESX)
		$Esxcli = Get-Esxcli -VMHost $ESX.Name
		$Delta = [Convert]::ToInt64($esxcli.system.uuid.get.Invoke().Split('-')[0],16)
		(Get-Date -Year 1970 -Day 1 -Month 1 -Hour 0 -Minute 0 -Second 0).AddSeconds($delta)
	} -Force > $Null
	$InstallDate = $(Get-VMHost -Name $vmhost | Select-Object -ExpandProperty EsxInstallDate)
	
	If ($ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Etat_ESX).Text -eq "Connected") {
		If ($ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Balise).Text -Like "@*") {
			If ($oESX_NTP -eq $Gateway) {
				$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Serveurs_NTP).Font.ColorIndex = 1
				$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Serveurs_NTP).Font.Bold = $False }
			Else {
				$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Serveurs_NTP).Font.ColorIndex = $Excel_Couleur_Error
				$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Serveurs_NTP).Font.Bold = $True
			}
		}
	}
		
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Datacenter)			= "$DC"										# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Cluster)			= "$Cluster"								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Etat_ESX)			= $ESX_State								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_vCenter)			= $vCenter									# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Balise)				= "$oESXTAG"								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Statut_ESX)			= "PoweredOn"								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Version)			= $ESX.Version + " - Build " + $ESX.Build	# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Date_Installation)	= "'" + $InstallDate						# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Serveurs_NTP)		= "$oESX_NTP"								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Zone_Temps)			= $UTC_Var1 + "+" + $UTC_Var2				# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Heure_Locale)		= $ESXCLI.Hardware.Clock.Get.Invoke()		# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Niveau_Compliance)	= $ComplianceLevel							# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Etat_Demon_NTP)		= $NTPD_status								# Ligne, Colonne
	
	### Colorisation des cellules
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Datacenter).Interior.ColorIndex 		= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Cluster).Interior.ColorIndex			= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Etat_ESX).Interior.ColorIndex			= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_vCenter).Interior.ColorIndex			= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Statut_ESX).Interior.ColorIndex			= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Date_Installation).Interior.ColorIndex	= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Heure_Locale).Interior.ColorIndex		= $Excel_Couleur_Background
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de la configuration SAN
Function Get-ESX_SAN { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	
	### Initialisation des variables
	$Global:Target				= $Null
	$Global:Target_Count 		= $Null
	$Global:LUNs				= $Null
	$Global:nrPaths				= $Null
	$Global:PathsSum			= $Null
	$Global:LUNsSum				= $Null
	$Global:PathsDead_Total		= $Null
	$Global:PathsActive_Total	= $Null
	$Global:PathsRedondance		= $Null
	$Global:PathsDead			= $Null
	$Global:PathsActive			= $Null	
	
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "SAN`t`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs SAN")
	$ElapsedTime_Start = (Get-Date)
	
	Get-VMHostStorage -RescanAllHba -VMHost $ESX -Server $vCenter | Out-Null
	
	[String[]]$ArraySAN_HBAName			= @()	# Initialisation du tableau relatif au nom des interfaces HBA actives
	$ArraySAN_LUNs						= @()	# Initialisation du tableau relatif au nombre de périphériques SAN (LUN)
	$ArraySAN_Paths						= @()	# Initialisation du tableau relatif au nombre de chemin SAN
	[String[]]$ArraySAN_Paths_Active	= @()	# Initialisation du tableau relatif au nombre de chemin SAN "active"
	$ArraySAN_Paths_Active_INT			= @()	# Initialisation du tableau relatif au nombre de chemin SAN "active" (format 'ENTIER')
	[String[]]$ArraySAN_Paths_Dead		= @()	# Initialisation du tableau relatif au nombre de chemin SAN "dead"
	$ArraySAN_Paths_Dead_INT			= @()	# Initialisation du tableau relatif au nombre de chemin SAN "dead" (format 'CHAINE')
	
	$Target_Count = (Get-VMHostHba -VMHost $ESX.Name -Type "FibreChannel").Count
	
	ForEach($hba in (Get-VMHostHba -VMHost $ESX.Name -Type "FibreChannel")) {
		$Target = ((Get-View $hba.VMhost).Config.StorageDevice.ScsiTopology.Adapter | Where {$_.Adapter -eq $hba.Key}).Target
		If ($Target_Count -eq 0) { Continue }
		$LUNs = Get-ScsiLun -Hba $hba -LunType "disk" -ErrorAction SilentlyContinue
		$nrPaths = ($Target | %{$_.Lun.Count} | Measure-Object -Sum).Sum
		
		# Boucle déterminant les valeurs des variables pour chacune des interfaces HBA
		$ArraySAN_Line = (1..$Target_Count)
		ForEach($L in $ArraySAN_Line)	{
			$ArraySAN_HBAName	+= $hba.Name
			$ArraySAN_LUNs		+= $LUNS.Count
			$ArraySAN_Paths		+= $nrPaths
		}
		
		$ESXCLI = Get-Esxcli -VMHost $ESX -Server $vCenter
		$ArraySAN_Paths_Active		+= ($ESXCLI.storage.core.path.list.invoke() | Where {($_.State -eq "active") -and ($_.Plugin -eq 'PowerPath') -and ($_.Adapter -eq $hba.Name)}).Count
		$ArraySAN_Paths_Active_INT 	+= ($ESXCLI.storage.core.path.list.invoke() | Where {($_.State -eq "active") -and ($_.Plugin -eq 'PowerPath') -and ($_.Adapter -eq $hba.Name)}).Count
		
		$ArraySAN_Paths_Dead		+= ($ESXCLI.storage.core.path.list.invoke() | Where {($_.State -eq "dead") -and ($_.Plugin -eq 'PowerPath') -and ($_.Adapter -eq $hba.Name)}).Count
		$ArraySAN_Paths_Dead_INT	+= ($ESXCLI.storage.core.path.list.invoke() | Where {($_.State -eq "dead") -and ($_.Plugin -eq 'PowerPath') -and ($_.Adapter -eq $hba.Name)}).Count
    }

	If ($ArraySAN_LUNs[1] -eq $ArraySAN_LUNs[2] -and $ArraySAN_Paths[1] -eq $ArraySAN_Paths[2]) { $PathsRedondance = "OK" }
	Else { $PathsRedondance = "NOK" }
	$PathsSum 			= ($ArraySAN_Paths[1] + $ArraySAN_Paths[2])
	$LUNsSum 			= ($ArraySAN_LUNs[1] + $ArraySAN_LUNs[2])
	$PathsDead_Total 	= ($ArraySAN_Paths_Dead_INT[0] + $ArraySAN_Paths_Dead_INT[1])
	$PathsActive_Total 	= ($ArraySAN_Paths_Active_INT[0] + $ArraySAN_Paths_Active_INT[1])
	
	$Global:PathsDead 	= "Total: $PathsDead_Total" + $VBCrLF + "(" + $ArraySAN_HBAName[1] + ": " + $ArraySAN_Paths_Dead[0] + ")" + $VBCrLF + "(" + $ArraySAN_HBAName[2] + ": " + $ArraySAN_Paths_Dead[1] + ")"
	$Global:PathsActive = "Total: $PathsActive_Total" + $VBCrLF + "(" + $ArraySAN_HBAName[1] + ": " + $ArraySAN_Paths_Active[0] + ")" + $VBCrLF + "(" + $ArraySAN_HBAName[2] + ": " + $ArraySAN_Paths_Active[1] + ")"
	
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_SAN_Total_Chemins) 			= "$PathsSum" + " (" + $Target_Count + " HBA)"		# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_SAN_Total_LUNs)				= "$LUNsSum" + " (" + $Target_Count + " HBA)"		# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_SAN_Total_Chemins_MORTS) 	= $PathsDead										# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_SAN_Total_Chemins_ACTIFS)	= $PathsActive										# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_SAN_Redondance_Chemins)		= $PathsRedondance									# Ligne, Colonne
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de configuration ESXi
Function Get-ESX_HOST { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )

	### Initialisation des variables
	$Global:CurrentEVCMode	= $Null
	$Global:MaxEVCMode		= $Null
	$Global:CPUPerformance	= $Null
	$Global:HyperThreading	= $Null
	$Global:Alarm			= $Null
	$Global:TPS				= $Null
	$Global:LPages			= $Null
	$Global:HAEnabled		= $Null
	
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "HOTE`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs HOTE")
	$ElapsedTime_Start = (Get-Date)
	
	$Global:CurrentEVCMode	= 	(Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Summary.CurrentEVCModeKey
	$Global:MaxEVCMode		= 	(Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Summary.MaxEVCModeKey
	$Global:CPUPerformance	= 	(Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Hardware.CpuPowerManagementInfo.CurrentPolicy
	$Global:HyperThreading	= 	(Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Config.HyperThread.Active
	If ((Get-VmHostService -VMHost $ESX -Server $vCenter | Where-Object {$_.key -eq "vmware-fdm"}).Running -eq "True") { $HAEnabled = "Running" } Else { $HAEnabled = "Not Running" }
	$Global:Alarm 			= (Get-VMHost -Name $ESX).ExtensionData.AlarmActionsEnabled
	If ($Alarm -eq "True") { $Alarm = "Enabled" } Else { $Alarm = "Disabled" }
	If ((Get-AdvancedSetting -Entity $ESX -Name "Mem.AllocGuestLargePage").Value -eq "0") 	{ $LPages = "Disabled" } 	Else { $LPages = "Enabled" }
	If ((Get-AdvancedSetting -Entity $ESX -Name "Mem.ShareForceSalting").Value -eq "0") 	{ $TPS = "Enabled" } 		Else { $TPS = "Disabled" }

	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HOST_Current_CPU_Policy)		= $CPUPerformance											# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HOST_Mode_HA)				= $HAEnabled												# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HOST_Hyperthreading)			= $HyperThreading											# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HOST_EVC)					= "Current: " + $CurrentEVCMode #+ ", Max: " + $MaxEVCMode	# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HOST_Etat_Alarmes)			= $Alarm													# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HOST_TPS_Salting)			= $TPS														# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_HOST_Larges_Pages_RAM)		= $LPages													# Ligne, Colonne
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de configuration du RESEAU
Function Get-ESX_NETWORK { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )

	### Initialisation des variables
	$Global:ESXAdapter		= $Null
	$Global:ESXAdapterName	= $Null
	$Global:ESXAdapterCount	= $Null
	$vLANID_Count			= $Null
	$vMotionIP				= $Null
	$vMotionEnabled			= $Null
	
	
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "RESEAUX`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs RESEAUX")
	$ElapsedTime_Start = (Get-Date)
	
	$vLANID = @()
	ForEach($dvPortgroup in (Get-VirtualPortgroup -VMHost $ESX -Server $vCenter -Distributed)){ $vLANID += $dvPortgroup.ExtensionData.Config.DefaultPortConfig.Vlan.VlanId; $vLANID_Count += 1 }
	$vLANID = $vLANID | Sort-Object
	$vLANID = "(" + $vLANID.Count + ") " + $vLANID + " [Distributed]"
	
	### Vérification si le nombre de VLAN précédemment calculé n'est pas vide
	If ($vLANID_Count -le 2) {
		$vLANID = @()
		ForEach($vPortgroup in (Get-VirtualPortgroup -VMHost $ESX -Server $vCenter)){ $vLANID += $vPortgroup.VlanId }
		$vLANID = $vLANID | Sort-Object
		$vLANID = "(" + $vLANID.Count + ") " + $vLANID + " [Local]"
	}
	
	$ESXCLI = Get-ESXCLI -VMHost $ESX -Server $vCenter
	$Global:ESXAdapter 		= $ESXCLI.Network.nic.list.Invoke()	| Where {$_.Link -eq "Up"}
	$Global:ESXAdapterName 	= $ESXCLI.Network.nic.list.Invoke()	| Where {$_.Link -eq "Up"} | Select Name, Speed, Duplex | Out-String
	$Global:ESXAdapterName 	= $ESXAdapterName -Replace "-","" -Replace "Name","" -Replace "Speed","" -Replace "Duplex","" -Replace "`r`n","" -Replace " ","" -Replace "10000", " 10000 " -Replace "vm", "`r`nvm"
	$Global:ESXAdapterCount = ($ESXCLI.Network.nic.list.Invoke() 	| Where {$_.Link -eq "Up"}).Count
	
	$vMotionIP = Get-VMHostNetworkAdapter -VMHost $ESX | Where {$_.DeviceName -eq "vmk1"} | Select IP | Out-String
	$vMotionIP = $vMotionIP -Replace "-","" -Replace "IP","" -Replace "`n","" -Replace " ",""
	$vMotionEnabled = (Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Summary.Config.VmotionEnabled
	If ($vMotionEnabled -eq $True) { $vMotionEnabled = "vMotion enabled" } Else { $vMotionEnabled = "vMotion disabled" }
	
	### Colorisation de la cellule s'il n'y a pas d'@ IP de vMotion
	If ($ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Etat_ESX).Text -eq "Connected") {
		If ($ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Balise).Text -Like "@*") {
			If ($vMotionIP) {
				$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_NETWORK_vMotion).Font.ColorIndex = 1
				$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_NETWORK_vMotion).Font.Bold = $False }
			Else {
				$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_NETWORK_vMotion).Font.ColorIndex = $Excel_Couleur_Error
				$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_NETWORK_vMotion).Font.Bold = $True
				$vMotionIP = "NULL"
			}
		}
	}

	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_NETWORK_VLAN)		= $vLANID											# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_NETWORK_Adaptateurs) = "(" + $ESXAdapterCount + ") " + $ESXAdapterName	# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_NETWORK_vMotion)		= $vMotionIP + " (" + $vMotionEnabled + ")"			# Ligne, Colonne
	
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_NETWORK_vMotion).Interior.ColorIndex = $Excel_Couleur_Background
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de configuration iLO
Function Get-HP_ILO { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )

	### Initialisation des variables
	$iLOversion = $Null
	
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "ILO`t`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs ILO")
	$ElapsedTime_Start = (Get-Date)
	
	$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ILO_Version) = $iLOversion # Ligne, Colonne
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
}


### Fonction chargée de la comparaison/l'homogénéité des valeurs entre elles
Function Get-ESX_Compare_Full {
	Write-Host "Vérification de l'homogénéité des ESXi du cluster... " -NoNewLine
	Write-Host "En cours" -ForegroundColor Green
	Write-Host "RAPPEL: $Mode [ID $Mode_Var]" -ForegroundColor Red
	LogTrace ("Vérification de l'homogénéité des ESXi du cluster...")
	
	### Initialisation de la matrice 2 dimensions
	Write-Host " * Initialisation de la matrice `t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Initialisation de la matrice")
	$ElapsedTime_Start = (Get-Date)
	
	$Ref_ArrayTag 		= $Null
	$Ref_ArrayVersion	= $Null
	$Ref_ArrayTimeZone 	= $Null
	$Ref_ArrayBaseline 	= $Null
	$Ref_ArrayNTPd	 	= $Null
	$Ref_ArraySAN 		= $Null
	$Ref_ArrayLUN 		= $Null
	$Ref_ArraySANDeath	= $Null
	$Ref_ArraySANAlive	= $Null
	$Ref_ArrayCPUPolicy	= $Null
	$Ref_ArrayHA		= $Null
	$Ref_ArrayHyperTh	= $Null
	$Ref_ArrayEVC		= $Null
	$Ref_ArrayAlarme	= $Null
	$Ref_ArrayTPS 		= $Null
	$Ref_ArrayLPages	= $Null
	$Ref_ArrayVLAN 		= $Null
	$Ref_ArrayLAN 		= $Null

	$ArrayTag 		= @()	# Initialisation du tableau relatif au TAG
	$ArrayVersion 	= @()	# Initialisation du tableau relatif à la version ESXi
	$ArrayTimeZone	= @()	# Initialisation du tableau relatif à la TimeZone
	$ArrayBaseline	= @()	# Initialisation du tableau relatif à la Baseline
	$ArrayNTPd 		= @()	# Initialisation du tableau relatif au démon NTP
	$ArraySAN 		= @()	# Initialisation du tableau relatif au nombre de chemins SAN
	$ArrayLUN 		= @()	# Initialisation du tableau relatif au nombre de LUN
	$ArraySANDeath	= @()	# Initialisation du tableau relatif au nombre de chemins SAN morts
	$ArraySANAlive	= @()	# Initialisation du tableau relatif au nombre de chemins SAN alive
	$ArrayCPUPolicy	= @()	# Initialisation du tableau relatif à l'utilisation CPU
	$ArrayHA		= @()	# Initialisation du tableau relatif au HA
	$ArrayHyperTh	= @()	# Initialisation du tableau relatif à l'HyperThreading
	$ArrayEVC		= @()	# Initialisation du tableau relatif au mode EVC
	$ArrayAlarme	= @()	# Initialisation du tableau relatif aux alarmes
	$ArrayTPS 		= @()	# Initialisation du tableau relatif au TPS
	$ArrayLPages	= @()	# Initialisation du tableau relatif aux Large Pages
	$ArrayVLAN 		= @()	# Initialisation du tableau relatif au nombre de VLAN ESX
	$ArrayLAN 		= @()	# Initialisation du tableau relatif au nombre d'adaptateurs LAN

	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime

	### Valorisation des tableaux relatifs aux colonnes à vérifier
	Write-Host " * Valorisation de la matrice `t`t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Valorisation de la matrice")
	$ElapsedTime_Start = (Get-Date)

	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
	LogTrace ("-- Traitement Excel de la ligne $ExcelLine_Start à la ligne $($ExcelLine_Start + $ESX_Counter - 1)")
	ForEach($L in $ExcelLine)	{
		### Si le nom du cluster existe dans la feuille de référence ET que l'ESXi n'est pas arrêté
		If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Statut_ESX).Text -ne "PoweredOff") {
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_ESX).Text -eq "Connected") {
				If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Text -Like "@*") {
					$ArrayTag		+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Text
					$ArrayVersion	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Text
					$ArrayTimeZone 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Text
					$ArrayBaseline 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Text
					$ArrayNTPd	 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Text
					$ArraySAN 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.IndexOf(" "))
					$ArrayLUN		+= $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Text
					$ArraySANDeath	+= $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Text.IndexOf("(") - 1)
					$ArraySANAlive	+= $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Text.IndexOf("(") - 1)
					$ArrayCPUPolicy	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Text
					$ArrayHA		+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Text
					$ArrayHyperTh	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Text
					$ArrayEVC		+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Text
					$ArrayAlarme	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Text
					$ArrayTPS 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Text
					$ArrayLPages 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Text
					$ArrayVLAN 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1)
					$ArrayLAN 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.IndexOf(")") - 1)
				}
			}
		}
	}
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime

	### Détermination des valeurs de références (majoritaires)
	Write-Host " * Détermination des valeurs de références`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Détermination des valeurs de références")
	$ElapsedTime_Start = (Get-Date)

	$Ref_ArrayTag 		= ($ArrayTag 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayVersion	= ($ArrayVersion 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayTimeZone 	= ($ArrayTimeZone 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayBaseline 	= ($ArrayBaseline 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayNTPd	 	= ($ArrayNTPd 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArraySAN 		= ($ArraySAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLUN 		= ($ArrayLUN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArraySANDeath	= ($ArraySANDeath	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArraySANAlive	= ($ArraySANAlive	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayCPUPolicy	= ($ArrayCPUPolicy 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHA		= ($ArrayHA 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHyperTh	= ($ArrayHyperTh 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayEVC		= ($ArrayEVC 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayAlarme	= ($ArrayAlarme 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayTPS 		= ($ArrayTPS		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLPages	= ($ArrayLPages 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayVLAN 		= ($ArrayVLAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLAN 		= ($ArrayLAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
	
	### Inscription des valeurs de références par cluster dans la 2ème feuille du fichier Excel
	Write-Host " * Inscription des valeurs de références dans Excel`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Inscription des valeurs de références dans Excel")
	$ElapsedTime_Start = (Get-Date)
	
	$GetName = $ExcelWorkSheet_Ref.Range("A1").EntireColumn.Find("")
	
	$ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_Cluster) = $Cluster.Name
	If ($Ref_ArrayVersion)		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVersion)		= $Ref_ArrayVersion		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVersion)		= "-" }
	If ($Ref_ArrayTimeZone) 	{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTimeZone) 	= $Ref_ArrayTimeZone 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTimeZone)		= "-" }
	If ($Ref_ArrayBaseline) 	{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayBaseline) 	= $Ref_ArrayBaseline 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayBaseline)		= "-" }
	If ($Ref_ArrayNTPd) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayNTPd) 		= $Ref_ArrayNTPd 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayNTPd)			= "-" }
	If ($Ref_ArraySAN) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArraySAN) 		= $Ref_ArraySAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArraySAN)			= "-" }
	If ($Ref_ArrayLUN) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLUN) 		= $Ref_ArrayLUN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLUN)			= "-" }
	If ($Ref_ArrayCPUPolicy)	{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayCPUPolicy)	= $Ref_ArrayCPUPolicy 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayCPUPolicy)	= "-" }
	If ($Ref_ArrayHA) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHA) 			= $Ref_ArrayHA 			} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHA)			= "-" }
	If ($Ref_ArrayHyperTh) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHyperTh) 	= $Ref_ArrayHyperTh 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHyperTh)		= "-" }
	If ($Ref_ArrayEVC) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayEVC) 		= $Ref_ArrayEVC 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayEVC)			= "-" }
	If ($Ref_ArrayAlarme) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayAlarme) 		= $Ref_ArrayAlarme 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayAlarme)		= "-" }
	If ($Ref_ArrayTPS) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTPS) 		= $Ref_ArrayTPS 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTPS)			= "-" }
	If ($Ref_ArrayLPages) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLPages)		= $Ref_ArrayLPages 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLPages)		= "-" }
	If ($Ref_ArrayVLAN) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVLAN) 		= $Ref_ArrayVLAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVLAN)			= "-" }
	If ($Ref_ArrayLAN) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLAN)			= $Ref_ArrayLAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLAN)			= "-" }
	
	Switch -Wildcard ($vCenter)	{
		"*VCSZ*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "PRODUCTION" }
		"*VCSY*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "NON PRODUCTION" }
		"*VCSQ*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "CLOUD" }
		"*VCSZY*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement) 	= "SITES ADMINISTRATIFS" }
		"*VCSSA*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "BAC A SABLE" }
		"*VCS00*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "PVM" }
		Default	{ $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)		= "NA" }
	}
	
	$ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime

	### Vérification des valeurs par rapport aux valeurs majoritaires
	Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Vérification des cellules vs valeurs majoritaires")
	$ElapsedTime_Start = (Get-Date)
	
	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ESX_Percent_NotCompliant_AVE = 0
	$ESX_NotCompliant_Item = 0
	
	$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
	LogTrace ("-- Traitement Excel de la ligne $ExcelLine_Start à la ligne $($ExcelLine_Start + $ESX_Counter - 1)")
	ForEach($L in $ExcelLine)	{
		$ESX_Percent_NotCompliant = 0
			
		### Si le nom du cluster existe dans la feuille de référence ET que l'ESXi n'est pas arrêté
		If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Statut_ESX).Text -ne "PoweredOff") {
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_ESX).Text -eq "Connected") {
				If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Text -Like "@*") {
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Serveurs_NTP).Font.ColorIndex -eq $Excel_Couleur_Error)	{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[Erreur NTP Serveur]" + $VBCrLF; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_vMotion).Font.ColorIndex -eq $Excel_Couleur_Error)	{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[Erreur @IP vMotion]" + $VBCrLF; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Text -ne $Ref_ArrayTag) 							{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Balise/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }									Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Font.ColorIndex = 1; 				$ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Text -ne $Ref_ArrayVersion) 					{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Version/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }							Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.ColorIndex = 1; 			$ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Text -ne $Ref_ArrayTimeZone) 				{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! TimeZone/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }					Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.ColorIndex = 1; 			$ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Text -ne $Ref_ArrayBaseline) 			{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Baseline/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }		Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = 1; 	$ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Text -ne $Ref_ArrayNTPd) 				{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! NTP démon/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }			Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = 1; 		$ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.IndexOf(" ")) -ne $Ref_ArraySAN) { $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details) = $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Chemins SAN/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Text -ne $Ref_ArrayLUN)						{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! LUNs/Ensemble]"; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 									Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.ColorIndex = 1; 			$ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Text.IndexOf("(") - 1) -ne $Ref_ArraySANDeath) { $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details) =$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[Chemin(s) mort(s)]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Text.IndexOf("(") - 1) -ne $Ref_ArraySANAlive) { $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details) =$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Chemin(s) présent(s)]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Redondance_Chemins).Text -ne "OK")			 			{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[Erreur redondance SAN]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Redondance_Chemins).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Redondance_Chemins).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }	Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Redondance_Chemins).Font.ColorIndex = 1; 	$ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Redondance_Chemins).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Text -ne $Ref_ArrayCPUPolicy)		{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[!CPU Policy/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }	Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = 1; 	$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Text -ne $Ref_ArrayHA) 							{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! HA/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }								Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.ColorIndex = 1; 			$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Text -ne $Ref_ArrayHyperTh) 				{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! HyperThreading/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }		Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.ColorIndex = 1; 		$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Text -ne $Ref_ArrayEVC) 							{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! EVC/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 										Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.ColorIndex = 1; 				$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Text -ne $Ref_ArrayAlarme)					{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Alarmes/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }					Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = 1; 		$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Text -ne $Ref_ArrayTPS) 					{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! TPS/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }						Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.ColorIndex = 1; 		$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Text -ne $Ref_ArrayLPages) 			{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Large Pages/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }		Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = 1; 	$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1) -ne $Ref_ArrayVLAN) { $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! VLANs/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.IndexOf(")") - 1) -ne $Ref_ArrayLAN) { $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details) = $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Adaptateurs LAN/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Font.Bold = $False }
				
					### Inscription du pourcentage de conformité
					$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Globale) = "[" + [math]::round((100 - (($ESX_Percent_NotCompliant / $ESX_Item_CheckType) * 100)), 0) + "% conforme]"
					$ESX_Percent_NotCompliant_AVE = [math]::round(($ESX_Percent_NotCompliant_AVE + (100 - (($ESX_Percent_NotCompliant / $ESX_Item_CheckType) * 100))) / 2, 2)
				} Else {
					$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Globale) = "-"
					$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details) = "ESXi en standby..."	}
			} Else {
				$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Globale) = "-"
				$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details) = "ESXi en maintenance..."	}
		} Else {
			$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Globale) = "-"
			$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details) = "ESXi OFF..."	}
		
		### Mise à jour de la colonne TimeStamp
		$ExcelWorkSheet.Cells.Item($L, $Excel_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
		
		### Valorisation des compteurs totaux
		If ($Excel_Conformite_Globale -eq "100%") {
			$ESX_Compliant += 1 }
		Else {
			$ESX_NotCompliant += 1
		}
		$ESX_NotCompliant_Total = $ESX_NotCompliant_Item
	}
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
	
	Write-Host "Vérification de l'homogénéité des données du cluster... " -NoNewLine
	Write-Host "Terminée`r`n" -ForegroundColor Black -BackgroundColor White
	Write-Host "--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Différence(s): $ESX_NotCompliant_Item`r`n * Taux moyen de conformité: $ESX_Percent_NotCompliant_AVE%`r`n------------------`r`n" -ForegroundColor White
	LogTrace ("Vérification de l'homogénéité des données du cluster... (Terminée)")
	LogTrace ("--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Différence(s): $ESX_NotCompliant_Item`r`n * Taux moyen de conformité: $ESX_Percent_NotCompliant_AVE%`r`n------------------`r`n")
}


### Fonction chargée de la comparaison/l'homogénéité des valeurs entre elles et relatives aux ESXi ciblés (Si le fichier REF n'existe pas)
Function Get-ESX_Compare_Cibles {
	Write-Host "Vérification de l'homogénéité des ESXi ciblés... " -NoNewLine
	Write-Host "En cours" -ForegroundColor Green
	Write-Host "RAPPEL: $Mode [ID $Mode_Var]" -ForegroundColor Red
	LogTrace ("Vérification de l'homogénéité des ESXi ciblés...")
	LogTrace ("RAPPEL: $Mode [ID $Mode_Var]")
	
	### Initialisation de la matrice 2 dimensions
	Write-Host " * Initialisation de la matrice `t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Initialisation de la matrice")
	$ElapsedTime_Start = (Get-Date)
	
	$Ref_ArrayTag 		= $Null
	$Ref_ArrayVersion	= $Null
	$Ref_ArrayTimeZone 	= $Null
	$Ref_ArrayBaseline 	= $Null
	$Ref_ArrayNTPd	 	= $Null
	$Ref_ArraySAN 		= $Null
	$Ref_ArrayLUN 		= $Null
	$Ref_ArraySANDeath	= $Null
	$Ref_ArraySANAlive	= $Null
	$Ref_ArrayCPUPolicy	= $Null
	$Ref_ArrayHA		= $Null
	$Ref_ArrayHyperTh	= $Null
	$Ref_ArrayEVC		= $Null
	$Ref_ArrayAlarme	= $Null
	$Ref_ArrayTPS 		= $Null
	$Ref_ArrayLPages	= $Null
	$Ref_ArrayVLAN 		= $Null
	$Ref_ArrayLAN 		= $Null

	$ArrayTag 		= @()	# Initialisation du tableau relatif au TAG
	$ArrayVersion 	= @()	# Initialisation du tableau relatif à la version ESXi
	$ArrayTimeZone	= @()	# Initialisation du tableau relatif à la TimeZone
	$ArrayBaseline	= @()	# Initialisation du tableau relatif à la Baseline
	$ArrayNTPd 		= @()	# Initialisation du tableau relatif au démon NTP
	$ArraySAN 		= @()	# Initialisation du tableau relatif au nombre de chemins SAN
	$ArrayLUN 		= @()	# Initialisation du tableau relatif au nombre de LUNs
	$ArraySANDeath	= @()	# Initialisation du tableau relatif au nombre de chemins SAN morts
	$ArraySANAlive	= @()	# Initialisation du tableau relatif au nombre de chemins SAN alive
	$ArrayCPUPolicy	= @()	# Initialisation du tableau relatif à l'utilisation CPU
	$ArrayHA		= @()	# Initialisation du tableau relatif au HA
	$ArrayHyperTh	= @()	# Initialisation du tableau relatif à l'HyperThreading
	$ArrayEVC		= @()	# Initialisation du tableau relatif au mode EVC
	$ArrayAlarme	= @()	# Initialisation du tableau relatif aux alarmes
	$ArrayTPS 		= @()	# Initialisation du tableau relatif au TPS
	$ArrayLPages	= @()	# Initialisation du tableau relatif aux Large Pages
	$ArrayVLAN 		= @()	# Initialisation du tableau relatif au nombre de VLAN ESX
	$ArrayLAN 		= @()	# Initialisation du tableau relatif au nombre d'adaptateurs LAN

	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime

	### Valorisation des tableaux relatifs aux colonnes à vérifier
	Write-Host " * Valorisation de la matrice `t`t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Valorisation de la matrice")
	$ElapsedTime_Start = (Get-Date)

	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
	LogTrace ("-- Traitement Excel de la ligne $ExcelLine_Start à la ligne $($ExcelLine_Start + $ESX_Counter - 1)")
	ForEach($L in $ExcelLine)	{
		### Si le nom du cluster existe dans la feuille de référence ET que l'ESXi n'est pas arrêté
		If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Statut_ESX).Text -ne "PoweredOff") {
			$ArrayTag		+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Text
			$ArrayVersion	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Text
			$ArrayTimeZone 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Text
			$ArrayBaseline 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Text
			$ArrayNTPd	 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Text
			$ArraySAN 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.IndexOf(" "))
			$ArrayLUN 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Text
			$ArraySANDeath	+= $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Text.IndexOf("(") - 1)
			$ArraySANAlive	+= $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Text.IndexOf("(") - 1)
			$ArrayCPUPolicy	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Text
			$ArrayHA		+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Text
			$ArrayHyperTh	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Text
			$ArrayEVC		+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Text
			$ArrayAlarme	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Text
			$ArrayTPS 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Text
			$ArrayLPages 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Text
			$ArrayVLAN 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1)
			$ArrayLAN 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.IndexOf(")") - 1)
		}
	}

	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime

	### Détermination des valeurs de références (majoritaires)
	Write-Host " * Détermination des valeurs de références`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Détermination des valeurs de références")
	$ElapsedTime_Start = (Get-Date)

	$Ref_ArrayTag 		= ($ArrayTag 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayVersion	= ($ArrayVersion 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayTimeZone 	= ($ArrayTimeZone 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayBaseline 	= ($ArrayBaseline 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayNTPd	 	= ($ArrayNTPd 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArraySAN 		= ($ArraySAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLUN 		= ($ArrayLUN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArraySANDeath	= ($ArraySANDeath	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArraySANAlive	= ($ArraySANAlive	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayCPUPolicy	= ($ArrayCPUPolicy 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHA		= ($ArrayHA 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHyperTh	= ($ArrayHyperTh 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayEVC		= ($ArrayEVC 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayAlarme	= ($ArrayAlarme 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayTPS 		= ($ArrayTPS		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLPages	= ($ArrayLPages 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayVLAN 		= ($ArrayVLAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLAN 		= ($ArrayLAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
	
	### Vérification des valeurs par rapport aux valeurs majoritaires
	Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Vérification des cellules vs valeurs majoritaires")
	$ElapsedTime_Start = (Get-Date)
	
	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ESX_Percent_NotCompliant_AVE = 0
	$ESX_NotCompliant_Item = 0
	
	$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
	LogTrace ("-- Traitement Excel de la ligne $ExcelLine_Start à la ligne $($ExcelLine_Start + $ESX_Counter - 1)")
	ForEach($L in $ExcelLine)	{
		$ESX_Percent_NotCompliant = 0
		
		### Si le nom du cluster existe dans la feuille de référence ET que l'ESXi n'est pas arrêté
		If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Statut_ESX).Text -ne "PoweredOff") {
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Serveurs_NTP).Font.ColorIndex -eq $Excel_Couleur_Error)	{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[Erreur NTP Serveur]" + $VBCrLF; 			$ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_vMotion).Font.ColorIndex -eq $Excel_Couleur_Error)	{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[Erreur @IP vMotion]" + $VBCrLF; 			$ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Text -ne $Ref_ArrayTag) 							{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Balise/Ensemble]" + $VBCrLF; 				$ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 							Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Balise).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Text -ne $Ref_ArrayVersion) 					{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Version/Ensemble]" + $VBCrLF; 			$ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }						Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Text -ne $Ref_ArrayTimeZone) 				{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! TimeZone/Ensemble]" + $VBCrLF; 			$ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }					Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Text -ne $Ref_ArrayBaseline) 			{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Baseline/Ensemble]" + $VBCrLF; 			$ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }	Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Text -ne $Ref_ArrayNTPd) 				{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! NTP démon/Ensemble]" + $VBCrLF; 			$ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }			Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.IndexOf(" ")) -ne $Ref_ArraySAN) { $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	=	$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Chemins SAN/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Text -ne $Ref_ArrayLUN) 						{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! LUNs/Ensemble]" + $VBCrLF;			 	$ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 					Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Text.IndexOf("(") - 1) -ne $Ref_ArraySANDeath) { $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details) =$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[Chemin(s) mort(s)]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_MORTS).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Text.IndexOf("(") - 1) -ne $Ref_ArraySANAlive) { $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details) =$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Chemin(s) présent(s)]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins_ACTIFS).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Redondance_Chemins).Text -ne "OK")			 			{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[Erreur redondance SAN]" + $VBCrLF; 		$ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Redondance_Chemins).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Redondance_Chemins).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 	Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Redondance_Chemins).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Redondance_Chemins).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Text -ne $Ref_ArrayCPUPolicy)		{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! CPU Policy/Ensemble]" + $VBCrLF; 		$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 	Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Text -ne $Ref_ArrayHA) 							{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! HA/Ensemble]" + $VBCrLF; 				$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 						Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Text -ne $Ref_ArrayHyperTh) 				{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! HyperThreading/Ensemble]" + $VBCrLF; 	$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 			Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Text -ne $Ref_ArrayEVC) 							{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! EVC/Ensemble]" + $VBCrLF; 				$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 								Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Text -ne $Ref_ArrayAlarme)					{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Alarmes/Ensemble]" + $VBCrLF; 			$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 				Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Text -ne $Ref_ArrayTPS) 					{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! TPS/Ensemble]" + $VBCrLF; 				$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 				Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Text -ne $Ref_ArrayLPages) 			{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Large Pages/Ensemble]" + $VBCrLF; 		$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 		Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1) -ne $Ref_ArrayVLAN) { $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= 	$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! VLANs/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.IndexOf(")") - 1) -ne $Ref_ArrayLAN) { $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= 	$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Adaptateurs LAN/Ensemble]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Font.Bold = $False }
			
			### Inscription du pourcentage de conformité
			$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Globale)= "[" + [math]::round((100 - (($ESX_Percent_NotCompliant / $ESX_Item_CheckType) * 100)), 0) + "% conforme]"
			$ESX_Percent_NotCompliant_AVE = [math]::round(($ESX_Percent_NotCompliant_AVE + (100 - (($ESX_Percent_NotCompliant / $ESX_Item_CheckType) * 100))) / 2, 2)
		}
		Else {
			$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Globale) = "-"
			$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details) = "ESXi OFF..."
		}
		
		### Mise à jour de la colonne TimeStamp
		$ExcelWorkSheet.Cells.Item($L, $Excel_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
		
		### Valorisation des compteurs totaux
		If ($Excel_Conformite_Globale -eq "100%") {
			$ESX_Compliant += 1 }
		Else {
			$ESX_NotCompliant += 1
		}
		$ESX_NotCompliant_Total = $ESX_NotCompliant_Item
	}
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
	
	Write-Host "Vérification de l'homogénéité des données ESXi ciblés...`t" -NoNewLine
	Write-Host "Terminée`r`n" -ForegroundColor Black -BackgroundColor White
	Write-Host "--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Différence(s): $ESX_NotCompliant_Item`r`n * Taux moyen de conformité: $ESX_Percent_NotCompliant_AVE%`r`n------------------`r`n" -ForegroundColor White
	LogTrace ("Vérification de l'homogénéité des données du cluster... (Terminée)")
	Logtrace ("--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Différence(s): $ESX_NotCompliant_Item`r`n * Taux moyen de conformité: $ESX_Percent_NotCompliant_AVE%`r`n------------------`r`n")
}


### Fonction chargée de la comparaison/l'homogénéité des valeurs avec celles du fichier de référence
Function Get-ESX_Compare_CiblesvsReference {

	### Initialisation des variables de références
	$Ref_ArrayVersion	= $Null
	$Ref_ArrayTimeZone 	= $Null
	$Ref_ArrayBaseline 	= $Null
	$Ref_ArrayNTPd	 	= $Null
	$Ref_ArraySAN		= $Null
	$Ref_ArrayLUN		= $Null
	$Ref_ArrayCPUPolicy	= $Null
	$Ref_ArrayHA		= $Null
	$Ref_ArrayHyperTh	= $Null
	$Ref_ArrayEVC		= $Null
	$Ref_ArrayAlarme	= $Null
	$Ref_ArrayTPS 		= $Null
	$Ref_ArrayLPages 	= $Null
	$Ref_ArrayVLAN		= $Null
	$Ref_ArrayLAN		= $Null

	### Recherche du cluster relatif à l'ESXi ciblé dans la feuille de référence Excel
	$GetName = $ExcelWorkSheet_Ref.Range("A1").EntireColumn.Find("$Cluster")
	
	### Si le cluster n'existe pas dans la feuille de référence
	If (!($GetName.Row)) {
		
		### Si les paramètres du fichier .BAT laissent supposer qu'il s'agit d'un nouveau cluster
		If ($ESX_Inc -eq "TOUS" -and $ESX_Exc -eq "AUCUN") {
			LogTrace ("-- Appel de la fonction 'Get-ESX_Compare_Cibles'")
			Get-ESX_Compare_Cibles
			
			Write-Host "Le fichier de référence va être mis à jour avec les nouvelles valeurs..." -ForegroundColor Yellow
			LogTrace ("Le fichier de référence va être mis à jour avec les nouvelles valeurs...")
			
			LogTrace ("-- Appel de la fonction 'Get-Ajout_Nouveau_Cluster'")
			Get-Ajout_Nouveau_Cluster
		}
		Else {
			Write-Host "Comparaison des ESXi ciblés impossible, cluster non trouvé... (Terminée)`r`n"
			LogTrace ("Comparaison des ESXi ciblés impossible, cluster non trouvé... (Terminée)")
			
			LogTrace ("-- Appel de la fonction 'Get-ESX_Compare_Cibles'")
			Get-ESX_Compare_Cibles
		}
	}
	
	### Si le cluster existe dans la feuille de référence
	Else {
		### Si les paramètres du fichier .BAT laissent supposer qu'il s'agit d'un nouveau cluster
		If ($ESX_Inc -eq "TOUS" -and $ESX_Exc -eq "AUCUN") {
			LogTrace ("-- Appel de la fonction 'Get-ESX_Compare_Cibles'")
			Get-ESX_Compare_Cibles
			
			Write-Host "Le fichier de référence va être mis à jour avec les nouvelles valeurs..." -ForegroundColor Yellow
			LogTrace ("Le fichier de référence va être mis à jour avec les nouvelles valeurs...")
			
			LogTrace ("-- Appel de la fonction 'Get-Ajout_Nouveau_Cluster'")
			Get-Ajout_Nouveau_Cluster
		}
		Else {
			Write-Host "Comparaison des ESXi ciblés avec les références clusters... " -NoNewLine
			Write-Host "En cours" -ForegroundColor Green
			LogTrace ("Comparaison des ESXi ciblés avec les références clusters...")
			
			### Valorisation des tableaux relatifs aux colonnes à vérifier
			Write-Host " * Valorisation des variables de références`t`t" -NoNewLine -ForegroundColor DarkYellow
			LogTrace (" * Valorisation des variables de références")
			$ElapsedTime_Start = (Get-Date)
			
			### Si le nom du cluster existe dans la feuille de référence ET que l'ESXi n'est pas arrêté
			If ($GetName.Row -and $ExcelWorkSheet.Cells.Item($GetName.Row, $Excel_CONF_Statut_ESX).Text -ne "PoweredOff") {
				$Ref_ArrayVersion	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVersion).Text	# Ligne, Colonne
				$Ref_ArrayTimeZone 	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTimeZone).Text	# Ligne, Colonne
				$Ref_ArrayBaseline 	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayBaseline).Text	# Ligne, Colonne
				$Ref_ArrayNTPd	 	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayNTPd).Text		# Ligne, Colonne
				$Ref_ArraySAN		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArraySAN).Text		# Ligne, Colonne
				$Ref_ArrayLUN		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLUN).Text		# Ligne, Colonne
				$Ref_ArrayCPUPolicy	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayCPUPolicy).Text	# Ligne, Colonne
				$Ref_ArrayHA		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHA).Text			# Ligne, Colonne
				$Ref_ArrayHyperTh	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHyperTh).Text	# Ligne, Colonne
				$Ref_ArrayEVC		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayEVC).Text		# Ligne, Colonne
				$Ref_ArrayAlarme	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayAlarme).Text		# Ligne, Colonne
				$Ref_ArrayTPS 		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTPS).Text		# Ligne, Colonne
				$Ref_ArrayLPages 	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLPages).Text		# Ligne, Colonne
				$Ref_ArrayVLAN		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVLAN).Text		# Ligne, Colonne
				$Ref_ArrayLAN		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLAN).Text		# Ligne, Colonne
			}
			
			LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
			Get-ElapsedTime

			### Vérification des données par rapport aux valeurs majoritaires
			Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
			LogTrace (" * Vérification des cellules vs valeurs majoritaires")
			$ElapsedTime_Start = (Get-Date)
				
			# Boucle déterminant la première ligne Excel du cluster à la dernière
			$ESX_Percent_NotCompliant_AVE = 0
			$ESX_NotCompliant_Item = 0

			$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
			LogTrace ("-- Traitement Excel de la ligne $ExcelLine_Start à la ligne $($ExcelLine_Start + $ESX_Counter - 1)")
			ForEach($L in $ExcelLine)	{
				$ESX_Percent_NotCompliant = 0
				
				### Si le nom du cluster existe dans la feuille de référence ET que l'ESXi n'est pas arrêté
				If ($GetName.Row -and $ExcelWorkSheet.Cells.Item($GetName.Row, $Excel_CONF_Statut_ESX).Text -ne "PoweredOff") {
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Serveurs_NTP).Font.ColorIndex -eq $Excel_Couleur_Error)	{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[Erreur NTP Serveur]" + $VBCrLF; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_vMotion).Font.ColorIndex -eq $Excel_Couleur_Error)	{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[Erreur @IP vMotion]" + $VBCrLF; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Text -ne $Ref_ArrayVersion) 					{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Version/Référence]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }							Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Text -ne $Ref_ArrayTimeZone)					{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! TimeZone/Référence]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }					Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Text -ne $Ref_ArrayBaseline) 			{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Baseline/Référence]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }		Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Text -ne $Ref_ArrayNTPd) 				{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! NTP démon/Référence]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }			Else { $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.IndexOf(" ")) -ne $Ref_ArraySAN)		{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Chemins SAN/Référence]" + $VBCrLF; 	$ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }	Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Text -ne $Ref_ArrayLUN)						{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! LUNs/Référence]" + $VBCrLF; 	$ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }						Else { $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Text -ne $Ref_ArrayCPUPolicy)		{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! CPU Policy/Référence]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1}	Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Text -ne $Ref_ArrayHA) 							{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! HA/Référence]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 								Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Text -ne $Ref_ArrayHyperTh) 				{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! HyperThreading/Référence]" + $VBCrLF;$ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }		Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Text -ne $Ref_ArrayEVC) 							{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! EVC/Référence]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 										Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Text -ne $Ref_ArrayAlarme)					{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Alarmes/Référence]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }				Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Text -ne $Ref_ArrayTPS) 					{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! TPS/Référence]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }						Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Text -ne $Ref_ArrayLPages) 			{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Large Pages/Référence]" + $VBCrLF; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1}		Else { $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1) -ne $Ref_ArrayVLAN) 	{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! VLAN/Référence]" + $VBCrLF; 			$ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }	Else { $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.IndexOf(")") - 1) -ne $Ref_ArrayLAN) 	{ $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Text + "[! Adaptateurs LAN/Référence]" + $VBCrLF; 			$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }	Else { $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details).Font.Bold = $False }

					### Inscription du pourcentage de conformité
					$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Globale) = "[" + [math]::round((100 - (($ESX_Percent_NotCompliant / 17) * 100)), 0) + "% conforme]"
					$ESX_Percent_NotCompliant_AVE = [math]::round(($ESX_Percent_NotCompliant_AVE + (100 - (($ESX_Percent_NotCompliant / 17) * 100))) / 2, 2)
				}
				Else {
					$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Globale) = "-"
					$ExcelWorkSheet.Cells.Item($L, $Excel_Conformite_Details) = "-"
				}
				
				### Mise à jour de la colonne TimeStamp
				$ExcelWorkSheet.Cells.Item($L, $Excel_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
				
				### Valorisation des compteurs totaux
				If ($Excel_Conformite_Globale -eq "100%") {
					$ESX_Compliant += 1 }
				Else {
					$ESX_NotCompliant += 1
				}
				$ESX_NotCompliant_Total = $ESX_NotCompliant_Item
			}

			LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
			Get-ElapsedTime

			Write-Host "Comparaison des ESXi ciblés avec les références clusters... (Terminée)`r`n"
			Write-Host "--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Cluster trouvé à la ligne: $($GetName.Row)`r`n * Différence(s): $ESX_NotCompliant_Item`r`n * Taux moyen de conformité: $ESX_Percent_NotCompliant_AVE%`r`n------------------`r`n" -ForegroundColor White
			LogTrace ("Comparaison des ESXi ciblés avec les références clusters...... (Terminée)")
			LogTrace ("--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Cluster trouvé à la ligne: $($GetName.Row)`r`n * Différence(s): $ESX_NotCompliant_Item`r`n * Taux moyen de conformité: $ESX_Percent_NotCompliant_AVE%`r`n------------------`r`n")
		}
	}
	
	$ExcelWorkBook_Ref.Save()
}


### Fonction chargée de la comparaison/l'homogénéité des valeurs entre elles
Function Get-Ajout_Nouveau_Cluster {

	Write-Host "Ajout/Mise à jour du cluster dans les références...`t" -NoNewLine
	Write-Host "En cours" -ForegroundColor Green
	LogTrace ("Ajout/Mise à jour du cluster dans les références...")

	### Initialisation de la matrice 2 dimensions
	Write-Host " * Initialisation de la matrice `t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Initialisation de la matrice")
	$ElapsedTime_Start = (Get-Date)
	
	$Ref_ArrayVersion	= $Null
	$Ref_ArrayTimeZone 	= $Null
	$Ref_ArrayBaseline 	= $Null
	$Ref_ArrayNTPd	 	= $Null
	$Ref_ArraySAN 		= $Null
	$Ref_ArrayLUN 		= $Null
	$Ref_ArrayCPUPolicy	= $Null
	$Ref_ArrayHA		= $Null
	$Ref_ArrayHyperTh	= $Null
	$Ref_ArrayEVC		= $Null
	$Ref_ArrayAlarme	= $Null
	$Ref_ArrayTPS 		= $Null
	$Ref_ArrayLPages	= $Null
	$Ref_ArrayVLAN 		= $Null
	$Ref_ArrayLAN 		= $Null
	
	$ArrayVersion 	= @()	# Initialisation du tableau relatif à la version ESXi
	$ArrayTimeZone	= @()	# Initialisation du tableau relatif à la TimeZone
	$ArrayBaseline	= @()	# Initialisation du tableau relatif à la Baseline
	$ArrayNTPd 		= @()	# Initialisation du tableau relatif au démon NTP
	$ArraySAN 		= @()	# Initialisation du tableau relatif au nombre de chemins SAN
	$ArrayLUN 		= @()	# Initialisation du tableau relatif au nombre de LUNs
	$ArrayCPUPolicy	= @()	# Initialisation du tableau relatif à l'utilisation CPU
	$ArrayHA		= @()	# Initialisation du tableau relatif au HA
	$ArrayHyperTh	= @()	# Initialisation du tableau relatif à l'HyperThreading
	$ArrayEVC		= @()	# Initialisation du tableau relatif au mode EVC
	$ArrayAlarme	= @()	# Initialisation du tableau relatif aux alarmes
	$ArrayTPS 		= @()	# Initialisation du tableau relatif au TPS
	$ArrayLPages	= @()	# Initialisation du tableau relatif aux Large Pages
	$ArrayVLAN 		= @()	# Initialisation du tableau relatif au nombre de VLAN ESXi
	$ArrayLAN 		= @()	# Initialisation du tableau relatif au nombre d'adaptateurs LAN

	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime

	### Valorisation des tableaux relatifs aux colonnes à vérifier
	Write-Host " * Valorisation de la matrice `t`t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Valorisation de la matrice")
	$ElapsedTime_Start = (Get-Date)
	
	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
	LogTrace ("-- Traitement Excel de la ligne $ExcelLine_Start à la ligne $($ExcelLine_Start + $ESX_Counter - 1)")
	ForEach($L in $ExcelLine)	{
		$ArrayVersion	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Version).Text
		$ArrayTimeZone 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Zone_Temps).Text
		$ArrayBaseline 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Niveau_Compliance).Text
		$ArrayNTPd	 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_CONF_Etat_Demon_NTP).Text
		$ArraySAN 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_Chemins).Text.IndexOf(" "))
		$ArrayLUN 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_SAN_Total_LUNs).Text
		$ArrayCPUPolicy	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Current_CPU_Policy).Text
		$ArrayHA		+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Mode_HA).Text
		$ArrayHyperTh	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Hyperthreading).Text
		$ArrayEVC		+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_EVC).Text
		$ArrayAlarme	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Etat_Alarmes).Text
		$ArrayTPS 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_TPS_Salting).Text
		$ArrayLPages 	+= $ExcelWorkSheet.Cells.Item($L, $Excel_HOST_Larges_Pages_RAM).Text
		$ArrayVLAN 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1)
		$ArrayLAN 		+= $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.Substring(1, $ExcelWorkSheet.Cells.Item($L, $Excel_NETWORK_Adaptateurs).Text.IndexOf(")") - 1)
	}
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime

	### Détermination des valeurs de références (majoritaires)
	Write-Host " * Détermination des valeurs de références`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Détermination des valeurs de références")
	$ElapsedTime_Start = (Get-Date)

	$Ref_ArrayVersion	= ($ArrayVersion 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayTimeZone 	= ($ArrayTimeZone 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayBaseline 	= ($ArrayBaseline 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayNTPd	 	= ($ArrayNTPd 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArraySAN 		= ($ArraySAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLUN 		= ($ArrayLUN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayCPUPolicy	= ($ArrayCPUPolicy 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHA		= ($ArrayHA 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHyperTh	= ($ArrayHyperTh 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayEVC		= ($ArrayEVC 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayAlarme	= ($ArrayAlarme 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayTPS 		= ($ArrayTPS		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLPages	= ($ArrayLPages 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayVLAN 		= ($ArrayVLAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLAN 		= ($ArrayLAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
	
	### Inscription des valeurs de références par cluster dans la 2ème feuille du fichier Excel
	Write-Host " * Inscription des valeurs de références dans Excel`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Inscription des valeurs de références dans Excel")
	$ElapsedTime_Start = (Get-Date)
	
	### Si les paramètres du fichier .BAT laissent supposer qu'il s'agit d'un nouveau cluster
	If ($ESX_Inc -eq "TOUS" -and $ESX_Exc -eq "AUCUN") {
		### Recherche d'une cellule vide dans la feuille de référence Excel
		$GetName = $ExcelWorkSheet_Ref.Range("A1").EntireColumn.Find("$Cluster")
		
		### Si le nom du cluster existe dans la feuille de référence
		If ($GetName.Row) {
			$ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_Cluster) = $Cluster.Name
			If ($Ref_ArrayVersion)		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVersion)		= $Ref_ArrayVersion		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVersion)		= "-" }
			If ($Ref_ArrayTimeZone) 	{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTimeZone) 	= $Ref_ArrayTimeZone 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTimeZone)		= "-" }
			If ($Ref_ArrayBaseline) 	{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayBaseline) 	= $Ref_ArrayBaseline 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayBaseline)		= "-" }
			If ($Ref_ArrayNTPd) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayNTPd) 		= $Ref_ArrayNTPd 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayNTPd)			= "-" }
			If ($Ref_ArraySAN) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArraySAN) 		= $Ref_ArraySAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArraySAN)			= "-" }
			If ($Ref_ArrayLUN) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLUN) 		= $Ref_ArrayLUN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLUN)			= "-" }
			If ($Ref_ArrayCPUPolicy)	{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayCPUPolicy)	= $Ref_ArrayCPUPolicy 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayCPUPolicy)	= "-" }
			If ($Ref_ArrayHA) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHA) 			= $Ref_ArrayHA 			} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHA)			= "-" }
			If ($Ref_ArrayHyperTh) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHyperTh) 	= $Ref_ArrayHyperTh 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHyperTh)		= "-" }
			If ($Ref_ArrayEVC) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayEVC) 		= $Ref_ArrayEVC 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayEVC)			= "-" }
			If ($Ref_ArrayAlarme) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayAlarme) 		= $Ref_ArrayAlarme 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayAlarme)		= "-" }
			If ($Ref_ArrayTPS) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTPS) 		= $Ref_ArrayTPS 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTPS)			= "-" }
			If ($Ref_ArrayLPages) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLPages)		= $Ref_ArrayLPages 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLPages)		= "-" }
			If ($Ref_ArrayVLAN) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVLAN) 		= $Ref_ArrayVLAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVLAN)			= "-" }
			If ($Ref_ArrayLAN) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLAN)			= $Ref_ArrayLAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLAN)			= "-" }
		}
		
		Switch -Wildcard ($vCenter)	{
			"*VCSZ*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "PRODUCTION" }
			"*VCSY*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "NON PRODUCTION" }
			"*VCSQ*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "CLOUD" }
			"*VCSZY*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement) 	= "SITES ADMINISTRATIFS" }
			"*VCSSA*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "BAC A SABLE" }
			"*VCS00*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "PVM" }
			Default	{ $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)		= "NA" }
		}
	}
	Else {
		### Recherche d'une cellule vide dans la feuille de référence Excel
		$GetName = $ExcelWorkSheet_Ref.Range("A1").EntireColumn.Find("")
		### Si le nom du cluster existe dans la feuille de référence
		If ($GetName.Row) {
			$ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_Cluster) = $Cluster.Name
			If ($Ref_ArrayVersion)		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVersion)		= $Ref_ArrayVersion		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVersion)		= "-" }
			If ($Ref_ArrayTimeZone) 	{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTimeZone) 	= $Ref_ArrayTimeZone 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTimeZone)		= "-" }
			If ($Ref_ArrayBaseline) 	{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayBaseline) 	= $Ref_ArrayBaseline 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayBaseline)		= "-" }
			If ($Ref_ArrayNTPd) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayNTPd) 		= $Ref_ArrayNTPd 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayNTPd)			= "-" }
			If ($Ref_ArraySAN) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArraySAN) 		= $Ref_ArraySAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArraySAN)			= "-" }
			If ($Ref_ArrayLUN) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLUN) 		= $Ref_ArrayLUN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLUN)			= "-" }
			If ($Ref_ArrayCPUPolicy)	{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayCPUPolicy)	= $Ref_ArrayCPUPolicy 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayCPUPolicy)	= "-" }
			If ($Ref_ArrayHA) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHA) 			= $Ref_ArrayHA 			} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHA)			= "-" }
			If ($Ref_ArrayHyperTh) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHyperTh) 	= $Ref_ArrayHyperTh 	} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayHyperTh)		= "-" }
			If ($Ref_ArrayEVC) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayEVC) 		= $Ref_ArrayEVC 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayEVC)			= "-" }
			If ($Ref_ArrayAlarme) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayAlarme) 		= $Ref_ArrayAlarme 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayAlarme)		= "-" }
			If ($Ref_ArrayTPS) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTPS) 		= $Ref_ArrayTPS 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayTPS)			= "-" }
			If ($Ref_ArrayLPages) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLPages)		= $Ref_ArrayLPages 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLPages)		= "-" }
			If ($Ref_ArrayVLAN) 		{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVLAN) 		= $Ref_ArrayVLAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayVLAN)			= "-" }
			If ($Ref_ArrayLAN) 			{ $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLAN)			= $Ref_ArrayLAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_ArrayLAN)			= "-" }
		}
					
		Switch -Wildcard ($vCenter)	{
			"*VCSZ*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "PRODUCTION" }
			"*VCSY*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "NON PRODUCTION" }
			"*VCSQ*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "CLOUD" }
			"*VCSZY*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement) 	= "SITES ADMINISTRATIFS" }
			"*VCSSA*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "BAC A SABLE" }
			"*VCS00*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "PVM" }
			Default	{ $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)		= "NA" }
		}
	}

	### Mise à jour de la colonne TimeStamp
	$ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
	
	$ExcelWorkBook_Ref.Save()
}


### Fonction chargée de la comparaison/l'homogénéité des clusters (En fin de traitement complet)
Function Get-Cluster_Compare {
	Write-Host "Vérification de l'homogénéité des clusters par ENV... " -NoNewLine
	Write-Host "En cours" -ForegroundColor Green
	LogTrace ("Vérification de l'homogénéité des clusters par ENV")
	
	### Récupération des valeurs majoritaires par cluster et par ENV
	Write-Host " * Initialisation de la matrice `t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Initialisation de la matrice")
	$ElapsedTime_Start = (Get-Date)
	
	$Ref_ArrayVersion	= $Null
	$Ref_ArrayTimeZone 	= $Null
	$Ref_ArrayBaseline 	= $Null
	$Ref_ArrayNTPd	 	= $Null
	$Ref_ArrayCPUPolicy	= $Null
	$Ref_ArrayHA		= $Null
	$Ref_ArrayHyperTh	= $Null
	$Ref_ArrayEVC		= $Null
	$Ref_ArrayAlarme	= $Null
	$Ref_ArrayTPS 		= $Null
	$Ref_ArrayLPages	= $Null
		
	$ArrayVersion 	= @()	# Initialisation du tableau relatif à la version ESXi
	$ArrayTimeZone	= @()	# Initialisation du tableau relatif à la TimeZone
	$ArrayBaseline	= @()	# Initialisation du tableau relatif à la Baseline
	$ArrayNTPd 		= @()	# Initialisation du tableau relatif au démon NTP
	$ArrayCPUPolicy	= @()	# Initialisation du tableau relatif à l'utilisation CPU
	$ArrayHA		= @()	# Initialisation du tableau relatif au HA
	$ArrayHyperTh	= @()	# Initialisation du tableau relatif à l'HyperThreading
	$ArrayEVC		= @()	# Initialisation du tableau relatif au mode EVC
	$ArrayAlarme	= @()	# Initialisation du tableau relatif aux alarmes
	$ArrayTPS 		= @()	# Initialisation du tableau relatif au TPS
	$ArrayLPages	= @()	# Initialisation du tableau relatif aux Large Pages
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
	
	### Valorisation des tableaux relatifs aux colonnes à vérifier
	Write-Host " * Valorisation de la matrice`t`t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Valorisation de la matrice")
	$ElapsedTime_Start = (Get-Date)
	
	### Recherche d'une cellule vide dans la feuille de référence Excel
	$GetName = $ExcelWorkSheet_Ref.Range("A1").EntireColumn.Find("")

	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = (2..($GetName.Row))
	LogTrace ("-- Traitement Excel de la ligne $ExcelLine_Start à la ligne $GetName.Row")
	ForEach($L in $ExcelLine)	{
		$ArrayVersion	+= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayVersion).Text
		$ArrayTimeZone 	+= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTimeZone).Text
		$ArrayBaseline 	+= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayBaseline).Text
		$ArrayNTPd	 	+= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayNTPd).Text
		$ArrayCPUPolicy	+= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayCPUPolicy).Text
		$ArrayHA		+= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHA).Text
		$ArrayHyperTh	+= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHyperTh).Text
		$ArrayEVC		+= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayEVC).Text
		$ArrayAlarme	+= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayAlarme).Text
		$ArrayTPS 		+= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTPS).Text
		$ArrayLPages 	+= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayLPages).Text
		
		### Colorisation des cellules
		$ExcelWorkSheet.Cells.Item($L, $Excel_Ref_Cluster).Interior.ColorIndex 			= $Excel_Couleur_Background
		$ExcelWorkSheet.Cells.Item($L, $Excel_Ref_ENVironnement).Interior.ColorIndex	= $Excel_Couleur_Background
		$ExcelWorkSheet.Cells.Item($L, $Excel_Ref_ArraySAN).Interior.ColorIndex			= $Excel_Couleur_Background
		$ExcelWorkSheet.Cells.Item($L, $Excel_Ref_ArrayLUN).Interior.ColorIndex			= $Excel_Couleur_Background
		$ExcelWorkSheet.Cells.Item($L, $Excel_Ref_ArrayVLAN).Interior.ColorIndex		= $Excel_Couleur_Background
		$ExcelWorkSheet.Cells.Item($L, $Excel_Ref_ArrayLAN).Interior.ColorIndex			= $Excel_Couleur_Background
	}
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime

	### Détermination des valeurs de références (majoritaires)
	Write-Host " * Détermination des valeurs de références`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Détermination des valeurs de références")
	$ElapsedTime_Start = (Get-Date)

	$Ref_ArrayVersion	= ($ArrayVersion 	| Group | ?{ $_.Count -gt ($Cluster_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayTimeZone 	= ($ArrayTimeZone 	| Group | ?{ $_.Count -gt ($Cluster_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayBaseline 	= ($ArrayBaseline 	| Group | ?{ $_.Count -gt ($Cluster_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayNTPd	 	= ($ArrayNTPd 		| Group | ?{ $_.Count -gt ($Cluster_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayCPUPolicy	= ($ArrayCPUPolicy 	| Group | ?{ $_.Count -gt ($Cluster_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHA		= ($ArrayHA 		| Group | ?{ $_.Count -gt ($Cluster_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHyperTh	= ($ArrayHyperTh 	| Group | ?{ $_.Count -gt ($Cluster_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayEVC		= ($ArrayEVC 		| Group | ?{ $_.Count -gt ($Cluster_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayAlarme	= ($ArrayAlarme 	| Group | ?{ $_.Count -gt ($Cluster_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayTPS 		= ($ArrayTPS		| Group | ?{ $_.Count -gt ($Cluster_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLPages	= ($ArrayLPages 	| Group | ?{ $_.Count -gt ($Cluster_Counter/2) }).Values | Select -Last 1
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
	
	### Vérification des données par rapport aux valeurs majoritaires
	Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Vérification des cellules vs valeurs majoritaires")
	$ElapsedTime_Start = (Get-Date)
	
	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$Global:Cluster_NotCompliant_Item = 0
	$Global:Cluster_Percent_NotCompliant_AVE = 0
		
	$ExcelLine = (2..($GetName.Row))
	LogTrace ("-- Traitement Excel de la ligne $ExcelLine_Start à la ligne $GetName.Row")
	ForEach($L in $ExcelLine)	{
		$Cluster_Percent_NotCompliant = 0
		
		If ($ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayVersion).Text -ne $Ref_ArrayVersion) 	{ $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details).Text + "[! Version/Ensemble]" + $VBCrLF; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayVersion).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayVersion).Font.Bold = $True; $Cluster_NotCompliant_Item += 1; $Cluster_Percent_NotCompliant += 1 } 			Else { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayVersion).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayVersion).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTimeZone).Text -ne $Ref_ArrayTimeZone) 	{ $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details).Text + "[! TimeZone/Ensemble]" + $VBCrLF; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTimeZone).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTimeZone).Font.Bold = $True; $Cluster_NotCompliant_Item += 1; $Cluster_Percent_NotCompliant += 1 } 		Else { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTimeZone).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTimeZone).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayBaseline).Text -ne $Ref_ArrayBaseline) 	{ $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details).Text + "[! Baseline/Ensemble]" + $VBCrLF; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayBaseline).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayBaseline).Font.Bold = $True; $Cluster_NotCompliant_Item += 1; $Cluster_Percent_NotCompliant += 1 } 		Else { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayBaseline).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayBaseline).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayNTPd).Text -ne $Ref_ArrayNTPd) 			{ $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details).Text + "[! NTP démon/Ensemble]" + $VBCrLF; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayNTPd).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayNTPd).Font.Bold = $True; $Cluster_NotCompliant_Item += 1; $Cluster_Percent_NotCompliant += 1 } 				Else { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayNTPd).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayNTPd).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayCPUPolicy).Text -ne $Ref_ArrayCPUPolicy) { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details).Text + "[! CPU Policy/Ensemble]" + $VBCrLF; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayCPUPolicy).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayCPUPolicy).Font.Bold = $True; $Cluster_NotCompliant_Item += 1; $Cluster_Percent_NotCompliant += 1 }	Else { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayCPUPolicy).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayCPUPolicy).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHA).Text -ne $Ref_ArrayHA) 				{ $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details).Text + "[! HA/Ensemble]" + $VBCrLF; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHA).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHA).Font.Bold = $True; $Cluster_NotCompliant_Item += 1; $Cluster_Percent_NotCompliant += 1 } 							Else { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHA).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHA).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHyperTh).Text -ne $Ref_ArrayHyperTh) 	{ $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details).Text + "[! HyperThreading/Ensemble]" + $VBCrLF; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHyperTh).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHyperTh).Font.Bold = $True; $Cluster_NotCompliant_Item += 1; $Cluster_Percent_NotCompliant += 1 } 	Else { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHyperTh).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayHyperTh).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayEVC).Text -ne $Ref_ArrayEVC) 			{ $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details).Text + "[! EVC/Ensemble]" + $VBCrLF; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayEVC).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayEVC).Font.Bold = $True; $Cluster_NotCompliant_Item += 1; $Cluster_Percent_NotCompliant += 1 } 						Else { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayEVC).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayEVC).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayAlarme).Text -ne $Ref_ArrayAlarme)		{ $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details).Text + "[! Alarmes/Ensemble]" + $VBCrLF; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayAlarme).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayAlarme).Font.Bold = $True; $Cluster_NotCompliant_Item += 1; $Cluster_Percent_NotCompliant += 1 } 			Else { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayAlarme).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayAlarme).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTPS).Text -ne $Ref_ArrayTPS) 			{ $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details).Text + "[! TPS/Ensemble]" + $VBCrLF; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTPS).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTPS).Font.Bold = $True; $Cluster_NotCompliant_Item += 1; $Cluster_Percent_NotCompliant += 1 } 						Else { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTPS).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayTPS).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayLPages).Text -ne $Ref_ArrayLPages) 		{ $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_Conformite_Details).Text + "[! Large Pages/Ensemble]" + $VBCrLF; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayLPages).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayLPages).Font.Bold = $True; $Cluster_NotCompliant_Item += 1; $Cluster_Percent_NotCompliant += 1 } 		Else { $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayLPages).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_ArrayLPages).Font.Bold = $False }
	
		### Inscription du pourcentage de conformité
		$ExcelWorkSheet.Cells.Item($L, $Excel_Ref_Conformite_Globale)= "[" + [math]::round((100 - (($Cluster_Percent_NotCompliant / $Cluster_Item_CheckType) * 100)), 0) + "% conforme]"
		$Cluster_Percent_NotCompliant_AVE = [math]::round(($Cluster_Percent_NotCompliant_AVE + (100 - (($Cluster_Percent_NotCompliant / $Cluster_Item_CheckType) * 100))) / 2, 2)
	
		### Mise à jour de la colonne TimeStamp
		$ExcelWorkSheet_Ref.Cells.Item($L, $Excel_Ref_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
		
		### Valorisation des compteurs totaux
		If ($Excel_Ref_Conformite_Globale -eq "100%") {
			$Cluster_NotCompliant += 1 }
		Else {
			$Cluster_NotCompliant += 1
		}
	}
	
	LogTrace ("-- Appel de la fonction 'Get-ElapsedTime'")
	Get-ElapsedTime
	
	Write-Host "Vérification de l'homogénéité des clusters par ENV... (Terminée)`r`n"
	Write-Host "--- En résumé: ---`r`n * Différence(s): $Cluster_NotCompliant_Item`r`n * Taux moyen de conformité: $Cluster_Percent_NotCompliant_AVE%`r`n" -ForegroundColor White
	LogTrace ("Vérification de l'homogénéité des clusters par ENV... (Terminée)")
	Logtrace ("--- En résumé: ---`r`n * Différence(s): $Cluster_NotCompliant_Item`r`n * Taux moyen de conformité: $Cluster_Percent_NotCompliant_AVE%`r`n")

	$ExcelWorkBook_Ref.Save()
}


### Fonction chargée de la conversion du format XLS en CSV
Function ConvertXLS2CSV {

	Write-Host "Conversion du fichier XLS en CSV...`t`t" -NoNewLine
	$ExcelWorkBook.SaveAs($File_CSV, 6)
	Write-Host "Terminé" -ForegroundColor Black -BackgroundColor White
	LogTrace ("Conversion du fichier XLS en CSV...`t`tTerminé")
}


### Fonction chargée de l'envoi de messages électroniques (En fin de traitement)
Function Send_Mail {

	$Sender1 		= "Christophe.VIDEAU-ext@ca-ts.fr"
	$Sender2 		= #"MCO.Infra.OS.distribues@ca-ts.fr"
	$Sender3 		= #"MCO.Infra.OS.distribues@ca-ts.fr"
	$Sender4 		= #"MCO.Infra.OS.distribues@ca-ts.fr"
	$From 			= "[ESXi] compliance check <ESXiCompliance.report@ca-ts.fr>"
	$Subject 		= "[Conformit&eacute; ESXi] Compte-rendu operationnel {Conformit&eacute; infrastructures VMware}"
	$Body 			= "Bonjour,<BR><BR>Vous trouverez en pi&egrave;ce jointe le fichier LOG et la synth&egrave;se technique de l&rsquo;op&eacute;ration.<BR><BR><U>En bref:</U><BR>Mode: " + $Mode + "<BR>Nombre d'ESXi conformes:<B><I> " + $ESX_Compliant + "</I></B><BR>Nombre d'ESXi non conforme(s):<B><I> " + $ESX_NotCompliant + " (" + $ESX_NotCompliant_Item + " diff&eacute;rences depuis le début)</I></B><BR>Nombre de clusters non conforme(s):<B><I> " + $Cluster_NotCompliant + " (" + $Cluster_NotCompliant_Item + " diff&eacute;rences)</I></B><BR>Dur&eacute;e du traitement: <I>" + $($Min) + "min. " + $($Sec) + "sec.</I><BR><BR>Cordialement.<BR>L'&eacute;quipe d&eacute;veloppement (Contact: Christophe VIDEAU)<BR><BR>Ce message a &eacute;t&eacute; envoy&eacute; par un automate, ne pas y r&eacute;pondre."
	If ($Mode_Var -eq 1) {
		$Attachments 	= $FicLog, $FicRes, $File_CSV }
	Else {
		$Attachments 	= $FicLog, $FicRes }
	$SMTP			 = "muz10-e1smtp-IN-DC-INT.zres.ztech"
	Send-MailMessage -To $Sender1, $Sender2, $Sender3, $Sender4 -From $From -Subject $Subject -Body $Body -Attachments $Attachments -SmtpServer $SMTP -Priority High -BodyAsHTML
}


### Fonction chargée de journaliser les actions du script dans un fichier LOG
Function LogTrace ($Message){
	If ($Mode_Debug -eq 1) {
		$Message = (Get-Date -format G) + " " + $Message + " [Ligne de code:" + $Myinvocation.ScriptlineNumber + "]" }
	Else {
		$Message = (Get-Date -format G) + " " + $Message }
	
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Message -Append
}

### Test de l'existence d'un processus en cours
If (Test-Path ($RepLog + $ScriptName + $File_PROCESS_LOCK)) {
	[System.Windows.Forms.MessageBox]::Show("Attention: Le script est déjà en cours d'exécution dans un autre processus. Attendre la fin de ce processus avant de l'exécuter de nouveau. Fin du script...", "Avertissement" , 0, 48)
	Exit
}

### Création du répertoire de LOG si besoin
If (!(Test-Path $RepLog)) { New-Item -Path $PathScript -ItemType Directory -Name "LOG"}
If ($args[1]) { $FicLog = $RepLog + $ScriptName + "_" + $Format_DATE + $File_LOG_ONESHOT } Else { $FicLog = $RepLog + $ScriptName + $File_LOG_DISCOVER }
$LineSep = "=" * 70

### Si le fichier LOG n'existe pas on le crée à vide
$Line = ">> DEBUT script de contrôle ESXi <<"
If (!(Test-Path $FicLog)) {
	$Line = (Get-Date -format G) + " " + $Line
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Line }
Else { LogTrace ($Line) }

### Bouchon bordelais pour tests
#[System.Windows.Forms.MessageBox]::Show("Attention: Mode DEBUG activé...", "Avertissement" , 0, 48)
#$args = @()
#$args = "SWMUZV1VCSZD.zres.ztech", "AUCUN", "AUCUN", "CL_MU_HDI_Z11,CL_MU_HDM_Z11,CT_MU_SMH_Z18", "TOUS"

### Test du contenu des paramètres de la ligne de commandes
If ($args[1]) {
	$Mode		= "Mode 'Par vCenter avec filtrage'"
	$Mode_Var	= 2
	If ($args[4] -eq "TOUS" -and $args[2] -eq "AUCUN") {
		Write-Host "$Mode [ID $Mode_Var]" -NoNewLine -ForegroundColor Red -BackgroundColor White
		Write-Host " - " -NoNewLine
		Write-Host "Inclus: TOUS, Exclus: AUCUN" -ForegroundColor Red -BackgroundColor White
		LogTrace ("$Mode [ID $Mode_Var] - Inclus: TOUS, Exclus: AUCUN")
	} Else {
		Write-Host "$Mode [ID $Mode_Var]" -ForegroundColor Red -BackgroundColor White
		LogTrace ("$Mode [ID $Mode_Var]")
	}
	
	
	### Test de l'existence du fichier Excel de référence
	Add-Type -AssemblyName Microsoft.Office.Interop.Excel
	$Excel = New-Object -ComObject Excel.Application
	
	$FicRes		= $RepLog + $ScriptName + "_" + $Format_DATE + $File_XLSX_ONESHOT
	$FicRes_Ref	= $RepLog + $ScriptName + $File_XLSX_DISCOVER
	
	LogTrace ("Création du fichier LOCK")
	New-Item ($RepLog + $ScriptName + $File_PROCESS_LOCK) -ItemType File | Out-Null
	
	If (Test-Path ($RepLog + $ScriptName + $File_XLSX_DISCOVER)) {
		Write-Host "Le fichier Excel de référence '$($RepLog + $ScriptName + $File_XLSX_DISCOVER)' est disponible" -ForegroundColor Green
		Write-Host "Comparaison possible en fin de traitement de chacun des clusters..." -ForegroundColor White
		LogTrace ("Le fichier Excel de référence '$($RepLog + $ScriptName + $File_XLSX_DISCOVER)' est disponible. Comparaison possible en fin de traitement de chacun des clusters...")
		$ExcelWorkBook_Ref	= $Excel.WorkBooks.Open($FicRes_Ref)	# Ouverture du fichier _REF
		$ExcelWorkSheet_Ref	= $Excel.WorkSheets.item(2)				# Définition de la feuille Excel par défaut du fichier _REF
	} Else {
		Write-Host "INFO: Le fichier Excel de référence '$($RepLog + $ScriptName + $File_XLSX_DISCOVER)' est indisponible.`r`nComparaison impossible en fin de traitement de chacun des clusters..." -ForegroundColor Red
		LogTrace ("INFO: Le fichier Excel de référence '$($RepLog + $ScriptName + $File_XLSX_DISCOVER)' est indisponible.`r`nComparaison impossible en fin de traitement de chacun des clusters...") }
	
	Copy-Item ($PathScript + $File_TEMPLATE) -Destination ($RepLog + $ScriptName + "_" + $Format_DATE + $File_XLSX_ONESHOT)
	
	### Définition du fichier Excel
	$ExcelWorkBook		= $Excel.WorkBooks.Open($FicRes)		# Ouverture du fichier [ONE SHOT]
	$ExcelWorkSheet		= $Excel.WorkSheets.item(1)				# Définition de la feuille Excel par défaut du fichier [ONE SHOT]
	$Excel.WorkSheets.item(2).Delete()							# Suppression de la 2ème feuille (inutile) du fichier [ONE SHOT]
	
	$ExcelWorkSheet.Activate()
	$Excel.Visible		= $False
	
	$Global:vCenters	= $args[0] # Nom du vCenter
	$Global:Cluster_Exc = $args[1] # Clusters exclus
	$Global:ESX_Exc		= $args[2] # ESXi exclus
	$Global:Cluster_Inc	= $args[3] # Clusters inclus
	$Global:ESX_Inc		= $args[4] # ESXi inclus

	### Initialisation de la matrice des paramètres du fichier .BAT
	$Array_Cluster_Esx	= @()
	$Array_ESX_Exc 		= @()
	$Array_Cluster_Inc	= @()
	$Array_ESX_Inc		= @()
	
	$Cluster_Exc		= $Cluster_Exc.ToUpper().Trim()
	$ESX_Exc			= $ESX_Exc.ToUpper().Trim()
	$Cluster_Inc		= $Cluster_Inc.ToUpper().Trim()
	$ESX_Inc			= $ESX_Inc.ToUpper().Trim()

	If (($Cluster_Exc -eq "") -or ($Cluster_Exc -eq "NONE"))	{ $Cluster_Exc			= "AUCUN" }
	If (($ESX_Exc -eq "") -or ($ESX_Exc -eq "NONE"))			{ $ESX_Exc				= "AUCUN" }
	If ($Cluster_Exc -ne "AUCUN" )								{ $Array_Cluster_Esx	= $Cluster_Exc.split(",") }
	If ($ESX_Exc -ne "AUCUN" )									{ $Array_ESX_Exc		= $ESX_Exc.split(",") }
	If (($Cluster_Inc -eq "") -or ($Cluster_Inc -eq "NONE"))	{ $Cluster_Inc			= "TOUS" }
	If (($ESX_Inc -eq "") -or ($ESX_Inc -eq "NONE"))			{ $ESX_Inc				= "TOUS" }
	If ($Cluster_Inc -ne "TOUS" )								{ $Array_Cluster_Inc 	= $Cluster_Inc.split(","); If ($Cluster_Inc.Contains(",")) { $Array_Cluster_Inc_Counter = ($Cluster_Inc.split(",").GetUpperBound(0) + 1) } Else { $Array_Cluster_Inc_Counter = 1 } }
	If ($ESX_Inc -ne "TOUS" )									{ $Array_ESX_Inc 		= $ESX_Inc.split(","); If ($ESX_Inc.Contains(",")) { $Array_ESX_Inc_Counter = ($ESX_Inc.split(",").GetUpperBound(0) + 1) } Else { $Array_ESX_Inc_Counter = 1 } }

	LogTrace ("$Mode")
	LogTrace ("Mode DEBUG (0:Désactivé - 1:Activé) .. : $Mode_Debug")
	LogTrace ("Fichier liste des vCenter à traiter .. : $vCenters")
	Logtrace ("Cluster à exclure .................... : $Cluster_Exc")
	Logtrace ("ESXi à exclure ....................... : $ESX_Exc")
	Logtrace ("Cluster à prendre en compte .......... : $Cluster_Inc (" + $Array_Cluster_Inc_Counter + ")")
	Logtrace ("ESXi à prendre en compte ............. : $ESX_Inc (" + $Array_ESX_Inc_Counter + ")")
	LogTrace ($LineSep + $VBCrLF)
	$TabVcc = $vCenters.split(",")
} Else {
	$Mode = "Mode 'Par vCenter sans aucun filtrage'"
	$Mode_Var = 1
	Write-Host "$Mode [ID $Mode_Var]" -ForegroundColor Red -BackgroundColor White
	
	### Test de l'existence du fichier Excel de référence
	If (Test-Path ($RepLog + $ScriptName + $File_XLSX_DISCOVER)) {
		$Global:Reponse = [System.Windows.Forms.MessageBox]::Show("Attention: le fichier '$($RepLog + $File_XLSX_DISCOVER)' existe déjà, voulez-vous le remplacer en fin d'exécution ? ", "Confirmation" , 4, 32)
		If (($Global:Reponse -eq "Yes") -or ($Global:Reponse -eq "Oui")) {
			LogTrace ("Le fichier '$($RepLog + $ScriptName + $File_XLSX_DISCOVER)' sera écrasé en fin d'exécution...")
			If (Test-Path ($PathScript + $File_TEMPLATE)) {
				Copy-Item ($PathScript + $File_TEMPLATE) -Destination ($RepLog + $ScriptName + $File_XLSX_DISCOVER_work) }
			Else {
				[System.Windows.Forms.MessageBox]::Show("Attention: Le fichier modèle '$File_TEMPLATE' n'existe pas. Fin du script...", "Avertissement" , 0, 48)
				Exit }
		}
		Else	{
			Write-Host "Le fichier '$($RepLog + $ScriptName + $File_XLSX_DISCOVER)' n'a pas été écrasé. Fin du script..." -ForegroundColor Red
			LogTrace ("Le fichier '$($RepLog + $ScriptName + $File_XLSX_DISCOVER)' n'a pas été écrasé. Fin du script...")
			LogTrace ("FIN du script")
			Write-Host -NoNewLine "FIN du script. Appuyer sur une touche pour quitter...`r`n"
			$Null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
			Exit }
	} Else { Copy-Item ($PathScript + $File_TEMPLATE) -Destination ($RepLog + $ScriptName + $File_XLSX_DISCOVER_work) }
	
	$FicRes = $RepLog + $ScriptName + $File_XLSX_DISCOVER_work
	
	LogTrace ("Création du fichier LOCK")
	New-Item ($RepLog + $ScriptName + $File_PROCESS_LOCK) -ItemType File | Out-Null
	
	### Définition du fichier Excel
	$Excel = New-Object -ComObject Excel.Application
	$ExcelWorkBook	= $Excel.WorkBooks.Open($FicRes)
	$ExcelWorkSheet = $Excel.WorkSheets.item(1)
	$ExcelWorkSheet_Ref = $Excel.WorkSheets.item(2)
	$ExcelWorkSheet.Activate()
	$Excel.Visible	= $False
	
	$Global:vCenters = $args[0]		# Nom du vCenter
	LogTrace ("Mode 'Par vCenter sans aucun filtrage'")
	LogTrace ("Fichier liste des vCenter à traiter .. : $vCenters")
	LogTrace ($LineSep + $VBCrLF)
	$TabVcc = $vCenters.split(",") }

	
### Boucle de traitement des vCenters contenus dans les paramètres de la ligne de commandes
ForEach ($vCenter in $TabVcc) {
	LogTrace ("DEBUT du traitement VCENTER $vCenter")
	Write-Host "`r`nDEBUT du traitement VCENTER " -NoNewLine
	Write-Host "$vCenter... ".ToUpper() -ForegroundColor Yellow -NoNewLine
	Write-Host "En cours" -ForegroundColor Green	
	
	$rccnx = Connect-VIServer -Server $vCenter -Protocol https -Credential $Credential
	$topCnxVcc = "0"
	If ($rccnx -ne $Null) { If ($rccnx.Isconnected) { $topCnxVcc = "1" } }

	If ($topCnxVcc -ne "1") { LogTrace ("ERREUR: Connexion KO au vCenter $vCenter - Arrêt du script")
		Write-Host "ERREUR: Connexion KO au vCenter $vCenter - Arrêt du script" -ForegroundColor White -BackgroundColor Red	
		$rc += 1
		Exit $rc }
	Else { LogTrace ("SUCCES: Connexion OK au vCenter $vCenter" + $VBCrLF)
		Write-Host "SUCCES: Connexion OK au vCenter $vCenter" -ForegroundColor Black -BackgroundColor Green }
	
	$Global:noDatacenter = 0
	$Global:oDatacenters = Get-Datacenter | Sort Name
	$Global:Datacenter_Counter = $oDatacenters.Count
	
	### Boucle de traitement des Datacenter composant le vCenter
	ForEach($DC in $oDatacenters){ $noDatacenter += 1
		LogTrace ("Traitement DATACENTER $DC n°$noDatacenter sur $Datacenter_Counter" + $VBCrLF)
		Write-Host "`r`nTraitement DATACENTER [#$noDatacenter/$Datacenter_Counter] " -NoNewLine
		Write-Host "$DC... ".ToUpper() -ForegroundColor Yellow -NoNewLine
		Write-Host "En cours" -ForegroundColor Green	

		If ($Mode_Var -eq 1) {
			$Global:noCluster = 0
		}
		$Global:oClusters = Get-Cluster -Location $DC | Sort Name
		$Global:Cluster_Counter = $oClusters.Count

		### Boucle de traitement des Clusters composants le Datacenter dont ceux contenus dans les paramètres de la ligne de commandes
		ForEach($Cluster in Get-Cluster -Location $DC | Sort Name)	{
			If (($Mode_Var -eq 2) -and ($Cluster_Inc -ne "TOUS"))	{
				$ClusterNom = $Cluster.Name
				$Global:Cluster_Counter = $Array_Cluster_Inc_Counter
				If ($Array_Cluster_Esx -Contains $ClusterNom) {
					Logtrace (" * Exclusion du cluster '$ClusterNom'...")
					Write-Host " * Exclusion " -NoNewLine -ForegroundColor Red
					Write-Host "du cluster " -NoNewLine
					Write-Host "'$ClusterNom'..." -ForegroundColor DarkYellow
					
					Continue }
					
				If (($Array_Cluster_Inc.Length -ne 0) -and ($Array_Cluster_Inc -notContains $ClusterNom)) {
					Logtrace (" * Cluster '$ClusterNom' absent des clusters à traiter...")
					Write-Host " * Cluster " -NoNewLine
					Write-Host "'$ClusterNom' " -NoNewLine -ForegroundColor DarkYellow
					Write-Host "absent " -NoNewLine -ForegroundColor Red
					Write-Host "des clusters à traiter..."
					
					Continue }
			}

			$Cluster_ID += 1
			$noCluster += 1
			
			### Exception de valorisation de la variable selon le mode
			If ($Mode_Var -eq 1) {	### Découverte complète ESXi
				$Global:ExcelLine_Start += $ESX_Counter
				LogTrace ("-- Valeur de la variable 'ExcelLine_Start'=$Global:ExcelLine_Start")
			}
			Else {					### Ciblage ESXi
				$Global:ExcelLine_Start += $No_ESX
				LogTrace ("-- Valeur de la variable 'ExcelLine_Start'=$Global:ExcelLine_Start")
			}

			LogTrace ("Traitement CLUSTER '$Cluster' n°$noCluster sur $Cluster_Counter {$Cluster_ID}")
			Write-Host "`r`nTraitement CLUSTER [#$noCluster/$Cluster_Counter] " -NoNewLine
			Write-Host "{$Cluster_ID} " -ForegroundColor Red -NoNewLine
			Write-Host "'$Cluster'... ".ToUpper() -ForegroundColor Yellow -NoNewLine
			Write-Host "En cours" -ForegroundColor Green

			$Global:oESX = Get-vmHost -Location $Cluster
			$Global:no_ESX = 0
			
			$Global:ESX_Counter = $oESX.Count
			LogTrace ("-- Valeur de la variable 'ESX_Counter'=$Global:ESX_Counter")

			### Boucle de traitement des ESXi composants le cluster dont ceux contenus dans les paramètres de la ligne de commandes
			ForEach($ESX in Get-vmHost -Location $Cluster) {
				If (($Mode_Var -eq 2) -and ($ESX_Inc -ne "TOUS")) {
					$Global:ESX_Counter = $Array_ESX_Inc_Counter
					LogTrace ("-- Valeur de la variable 'Array_ESX_Inc_Counter'=$Array_ESX_Inc_Counter")
					If ($Array_ESX_Exc -Contains $ESX) {
						Logtrace (" * Exclusion de l'ESXi '$ESX'...")
						Write-Host " * Exclusion " -NoNewLine -ForegroundColor Red
						Write-Host "de l'ESXi " -NoNewLines
						Write-Host "'$ESX'..." -ForegroundColor DarkYellow
						$No_ESX += 1		### Incrémentation de la valeur du n° d'ESXi
						LogTrace ("-- Valeur de la variable 'no_ESX'=$No_ESX")
						
						Continue
					}
					If (($Array_ESX_Inc.Length -ne 0) -and ($Array_ESX_Inc -notContains $ESX)) {
						Logtrace (" * ESXi '$ESX' absent des ESXi à traiter...")
						Write-Host " * ESXi " -NoNewLine
						Write-Host "'$ESX' " -NoNewLine -ForegroundColor DarkYellow
						Write-Host "absent " -NoNewLine -ForegroundColor Red
						Write-Host "des ESXi à traiter..."
						$No_ESX += 1		### Incrémentation de la valeur du n° d'ESXi
						LogTrace ("-- Valeur de la variable 'no_ESX'=$No_ESX")
						
						Continue
					}
				}
 
				$No_ESX += 1		### N° ESXi dans le cluster
				LogTrace ("-- Valeur de la variable 'no_ESX'=$No_ESX")
				
				$No_INC_ESX += 1	### n° ESXi total (depuis le début de l'exécution du script)
				LogTrace ("-- Valeur de la variable 'no_inc_ESX'=$No_INC_ESX")

				$StartTime = Get-Date -Format HH:mm:ss
				LogTrace ("Traitement ESXi '$ESX' n°$No_ESX sur $($ESX_Counter) {$No_INC_ESX}")
				Write-Host "[$StartTime] Traitement ESXi [#$No_ESX/$($ESX_Counter)] " -NoNewLine
				Write-Host "{$No_INC_ESX} " -ForegroundColor Red -NoNewLine
				Write-Host "'$ESX'... ".ToUpper() -ForegroundColor Yellow -NoNewLine
				Write-Host "En cours" -ForegroundColor Green

				If ($ESX.PowerState -ne "PoweredOn") {
					For ($i = 1; $i -le $Excel_Conformite_Details; $i++) {
						$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $i) = "NA"
						$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $i).Interior.ColorIndex = $Excel_Couleur_Background
					}
					$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_Nom_ESX)			= $ESX.Name			# Ligne, Colonne
					$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Statut_ESX) = "PoweredOff" 		# Ligne, Colonne
					$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Datacenter)	= "$DC"				# Ligne, Colonne
					$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_Cluster)	= "$Cluster"		# Ligne, Colonne
					$ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_CONF_vCenter)	= "$vCenter"		# Ligne, Colonne

					Switch -Wildcard ($vCenter)	{
						"*VCSZ*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "PRODUCTION" }
						"*VCSY*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "NON PRODUCTION" }
						"*VCSQ*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "CLOUD" }
						"*VCSZY*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement) 	= "SITES ADMINISTRATIFS" }
						"*VCSSA*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "BAC A SABLE" }
						"*VCS00*" { $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)	= "PVM" }
						Default	{ $ExcelWorkSheet.Cells.Item($No_INC_ESX + 1, $Excel_ENVironnement)		= "NA" }
					}

					$EndTime = Get-Date -Format HH:mm:ss
					
					### Enregistrement des modifications Excel
					$ExcelWorkBook.Save()
					
					LogTrace ("MISE A JOUR des données pour l'ESXi '$ESX'. Zone [L$ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1) *L$($No_INC_ESX + 1)*]..." + $VBCrLF)
					Write-Host "[$EndTime] Mise à jour des données Excel " -NoNewLine
					Write-Host "Zone [L$ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1) *L$($No_INC_ESX + 1)*]... " -ForegroundColor Yellow -NoNewLine
					Write-Host "Terminée`r`n" -ForegroundColor Black -BackgroundColor White

					# Dans le cas d'une découverte complète du périmètre
					If ($Mode_Var -eq 1) {
						If ($No_ESX -eq $($oESX.Count)) {
							LogTrace ("-- Appel de la fonction 'Get-ESX_Compare_Full'")
							Get-ESX_Compare_Full
							$ExcelWorkBook.Save()
						}
					}
					
					### Sélection de la fonction selon le mode de départ
					If ($Mode_Var -eq 2) {
						If (Test-Path ($RepLog + $ScriptName + $File_XLSX_DISCOVER)) {	# Dans le cas d'une vérification de conformité unitaire si le fichier Excel de référence existe
							If ($No_ESX -eq $ESX_Counter) {
								LogTrace ("-- Appel de la fonction 'Get-ESX_Compare_CiblesvsReference'")
								Get-ESX_Compare_CiblesvsReference
								$ExcelWorkBook.Save()
								
								Break
							}
						}
						Else {	# Dans le cas d'une vérification de conformité unitaire si le fichier Excel de référence n'existe pas
							If ($No_ESX -eq $ESX_Counter) {
								LogTrace ("-- Appel de la fonction 'Get-ESX_Compare_Cibles'")
								Get-ESX_Compare_Cibles
								$ExcelWorkBook.Save()
								
								Break
							}
						}
					}
					Continue
				}
				
				### Exécution des fonctions de récupération des données ESX
				LogTrace ("-- Appel de la fonction 'Get-ESX_HARD'")
				Get-ESX_HARD -vmHost $ESX		# Récupération matérielle
				
				LogTrace ("-- Appel de la fonction 'Get-ESX_CONFIG'")
				Get-ESX_CONFIG -vmHost $ESX		# Récupération de la configuration matérielle
				
				LogTrace ("-- Appel de la fonction 'Get-ESX_SAN'")
				Get-ESX_SAN -vmHost $ESX		# Récupération des données stockage SAN
				
				LogTrace ("-- Appel de la fonction 'Get-ESX_HOST'")
				Get-ESX_HOST -vmHost $ESX		# Récupération de la configuration ESXi
				
				LogTrace ("-- Appel de la fonction 'Get-ESX_NETWORK'")
				Get-ESX_NETWORK -vmHost $ESX	# Récupération de la configuration réseaux
				
				LogTrace ("-- Appel de la fonction 'Get-HP_ILO'")
				Get-HP_ILO -vmHost $ESX			# Récupération de la configuration iLO
				
				### Enregistrement des modifications Excel
				$ExcelWorkBook.Save()

				$EndTime = Get-Date -Format HH:mm:ss
				LogTrace ("MISE A JOUR des données pour l'ESXi '$ESX'. Zone [L$ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1) *L$($No_INC_ESX + 1)*]..." + $VBCrLF)
				Write-Host "[$EndTime] Mise à jour des données Excel " -NoNewLine
				Write-Host "Zone [L$ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1) *L$($No_INC_ESX + 1)*]... " -ForegroundColor Yellow -NoNewLine
				Write-Host "Terminée`r`n" -ForegroundColor Black -BackgroundColor White

				# Dans le cas d'une découverte complète du périmètre
				If ($Mode_Var -eq 1) {
					If ($No_ESX -eq $($oESX.Count)) {
						LogTrace ("-- Appel de la fonction 'Get-ESX_Compare_Full'")
						Get-ESX_Compare_Full
						$ExcelWorkBook.Save()
					}
				}
				
				### Sélection de la fonction selon le mode de départ
				If ($Mode_Var -eq 2) {
					If (Test-Path ($RepLog + $ScriptName + $File_XLSX_DISCOVER)) {	# Dans le cas d'une vérification de conformité unitaire si le fichier Excel de référence existe
						If ($No_ESX -eq $ESX_Counter) {
							LogTrace ("-- Appel de la fonction 'Get-ESX_Compare_CiblesvsReference'")
							Get-ESX_Compare_CiblesvsReference
							$ExcelWorkBook.Save()
							Break
						}
					}
					Else {	# Dans le cas d'une vérification de conformité unitaire si le fichier Excel de référence n'existe pas
						If ($No_ESX -eq $ESX_Counter) {
							LogTrace ("-- Appel de la fonction 'Get-ESX_Compare_Cibles'")
							Get-ESX_Compare_Cibles
							$ExcelWorkBook.Save()
							Break
						}
					}
				}
			}
		}
	}
	
	### Exécution des fonctions de récupération des valeurs ESX
	If ($Mode_Var -eq 1) {
		LogTrace ("-- Appel de la fonction 'Get-Cluster_Compare'")
		Get-Cluster_Compare
	}
	
	LogTrace ("DECONNEXION et FIN du traitement depuis le VCENTER '$vCenter'`r`n")
	Disconnect-VIServer -Server $vCenter -Force -Confirm:$False
}

$ExcelWorkBook.Save()

Write-Host "Enregistrement final du classeur Excel [Terminé]"
LogTrace ("Enregistrement final du classeur Excel [Terminé]")

### Exécution de la conversion du fichier XLS en CSV uniquement en mode découverte
If ($Mode_Var -eq 1) {
	LogTrace ("-- Appel de la fonction 'ConvertXLS2CSV'")
	ConvertXLS2CSV
}

### FIN du programme Excel
$Excel.Quit()

Start-Sleep -s 5

### Exécution de l'envoi du mail
LogTrace ("-- Appel de la fonction 'Send_Mail'")
Send_Mail
Write-Host "Envoi du mail avec les fichiers LOG et XLSX [Terminé]..." -ForegroundColor White
LogTrace ("Envoi du mail avec les fichiers LOG et XLSX [Terminé]...")

### Sauvegarde du précédent fichier _REF.xlsx, _REF.log et _REF.csv
If (($Global:Reponse -eq "Yes") -or ($Global:Reponse -eq "Oui")) {
	Move-Item ($RepLog + $ScriptName + $File_XLSX_DISCOVER) ($RepLog + $ScriptName + $File_XLSX_DISCOVER_bckp)		# Sauvegarde du précédent fichier _REF.xlsx
	If (Test-Path ($RepLog + $ScriptName + $File_LOG_DISCOVER)) { Move-Item ($RepLog + $ScriptName + $File_LOG_DISCOVER) ($RepLog + $ScriptName + $File_LOG_DISCOVER_bckp) }	# Si le fichier existe, sauvegarde du précédent fichier _REF.log
	If (Test-Path ($RepLog + $ScriptName + $File_CSV_DISCOVER)) { Move-Item ($RepLog + $ScriptName + $File_CSV_DISCOVER) ($RepLog + $ScriptName + $File_CSV_DISCOVER_bckp)	}	# Si le fichier existe, sauvegarde du précédent fichier _REF.csv
	
	Write-Host "Le fichier '$($RepLog + $ScriptName + $File_XLSX_DISCOVER)' (XLSX + LOG + CSV) a été sauvegardé" -ForegroundColor Red
	LogTrace ("Le fichier '$($RepLog + $ScriptName + $File_XLSX_DISCOVER)' (XLSX + LOG + CSV) a été sauvegardé")
}
	
### Renommage du fichier _REF-work vers le fichier _REF uniquement en mode découverte
If ($Mode_Var -eq 1) {
	Move-Item ($RepLog + $ScriptName + $File_XLSX_DISCOVER_work) ($RepLog + $ScriptName + $File_XLSX_DISCOVER)
	Move-Item ($RepLog + $ScriptName + $File_LOG_DISCOVER_work) ($RepLog + $ScriptName + $File_LOG_DISCOVER)
}

LogTrace ("Suppression du fichier LOCK")
Remove-Item -Path ($RepLog + $ScriptName + $File_PROCESS_LOCK) -Force

Write-Host -NoNewLine "FIN du script. Appuyer sur une touche pour quitter...`r`n"
$Null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

### FIN du script