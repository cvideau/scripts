### V0.4 (novembre 2018) - Développeur: Christophe VIDEAU
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
Clear-Host

Write-Host "Développement (2018) Christophe VIDEAU - Version 0.4`r`n" -ForegroundColor White

$Global:ElapsedTime_Start = 0; $Global:ElapsedTime_End = 0; $Global:ESX_Compliant = 0; $Global:ESX_NotCompliant = 0
$Global:no_inc_ESX = 0; $Global:ESX_Counter = 0; $Global:Cluster_Counter = 0; $Global:Cluster_ID = 0
$Global:Cluster = $NULL
$Global:ExcelLine_Start = 2		### Démarrage à la 2ème ligne du fichier Excel

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
$Excel_Couleur_Error		= 3
$Excel_Couleur_Background	= 15


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
	
	$Esxcli2 = Get-ESXCLI -VMHost $ESX.Name -Server $vCenter
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_Nom_ESX)	= $ESX.Name

	Switch -Wildcard ($vCenter)	{
		"*VCSZ*" { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement) 	= "PRODUCTION" }
		"*VCSY*" { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement) 	= "NON PRODUCTION" }
		"*VCSQ*" { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement) 	= "CLOUD" }
		"*VCSZY*" { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement) 	= "SITES ADMINISTRATIFS" }
		"*VCSSA*" { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement) 	= "BAC A SABLE" }
		Default	{ $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement) 	= "NA" }
	}
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement).Interior.ColorIndex = $Excel_Couleur_Background
	
	
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Constructeur)	= $Esxcli2.Hardware.Platform.Get.Invoke().VendorName
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Modele)			= $Esxcli2.Hardware.Platform.Get.Invoke().ProductName
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Num_Serie)		= $Esxcli2.Hardware.Platform.Get.Invoke().SerialNumber
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Version_BIOS)	= $ESX.ExtensionData.Hardware.BiosInfo.BiosVersion
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Date_BIOS)		= $ESX.ExtensionData.Hardware.BiosInfo.ReleaseDate
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Carte_SD)		= $Esxcli2.Storage.Core.Path.List.Invoke().AdapterTransportDetails | Where { $_.Device -eq "mpx.vmhba32:C0:T0:L0" }
	
	### Colorisation des cellules
	$Esxcli2 = Get-ESXCLI -VMHost $ESX.Name -Server $vCenter
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_Nom_ESX).Interior.ColorIndex 			= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Constructeur).Interior.ColorIndex	= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Modele).Interior.ColorIndex 		= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Num_Serie).Interior.ColorIndex 		= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Version_BIOS).Interior.ColorIndex 	= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Date_BIOS).Interior.ColorIndex 		= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HARD_Carte_SD).Interior.ColorIndex 		= $Excel_Couleur_Background
		
	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de CONFIGURATION
Function Get-ESX_CONFIG { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "CONFIGURATION`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs CONFIGURATION")
	$ElapsedTime_Start = (Get-Date)
	
	$Esxcli2 = Get-ESXCLI -VMHost $ESX.Name -Server $vCenter
	$Global:oESXTAG = Get-VMHost -Name $ESX | Get-TagAssignment | Select Tag
	$Global:oESX_NTP = Get-VMHostNtpServer -VMHost $ESX.Name -Server $vCenter
	
	If ((Get-VmHostService -VMHost $ESX -Server $vCenter | Where-Object {$_.key -eq "ntpd"}).Running -eq "True") { $ntpd_status = "Running" } Else { $ntpd_status = "Not Running" }
	If ($Esxcli2.System.MaintenanceMode.Get.Invoke() -eq "Enabled") { $ESX_State = "Maintenance" } Else { $ESX_State = "Connected" }
	If ((Get-Compliance -Entity $ESX -Detailed -WarningAction "SilentlyContinue" | WHERE {$_.NotCompliantPatches -ne $NULL} | SELECT Status).Count -gt 0) { $ComplianceLevel = "Baseline - Not compliant" } Else { $ComplianceLevel = "Baseline - Compliant" }
	$Global:UTC = Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}
	$Global:UTC_Var1 = $UTC.Config.DateTimeInfo.TimeZone.Name
	$Global:UTC_Var2 = $UTC.Config.DateTimeInfo.TimeZone.GmtOffset
	$Global:Gateway = Get-VmHostNetwork -Host $ESX  -Server $vCenter | Select VMkernelGateway -ExpandProperty VirtualNic | Where { $_.VMotionEnabled } | Select -ExpandProperty VMkernelGateway -WarningAction "SilentlyContinue"
	
	New-VIProperty -Name EsxInstallDate -ObjectType VMHost -Value { Param($ESX)
		$Esxcli = Get-Esxcli -VMHost $ESX.Name
		$Delta = [Convert]::ToInt64($esxcli.system.uuid.get.Invoke().Split('-')[0],16)
		(Get-Date -Year 1970 -Day 1 -Month 1 -Hour 0 -Minute 0 -Second 0).AddSeconds($delta)
	} -Force > $NULL
	$InstallDate = $(Get-VMHost -Name $vmhost | Select-Object -ExpandProperty EsxInstallDate)
	
	If ($ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Etat_ESX).Text -eq "Connected") {
		If ($ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Balise).Text -Like "@*") {
			If ($oESX_NTP -eq $Gateway) { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Serveurs_NTP).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Serveurs_NTP).Font.Bold = $False } Else { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Serveurs_NTP).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Serveurs_NTP).Font.Bold = $True }
		}
	}
		
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Datacenter)			= "$DC"										# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Cluster)			= "$Cluster"								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Etat_ESX)			= $ESX_State								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_vCenter)			= $vCenter									# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Balise)				= "$oESXTAG"								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Statut_ESX)			= "PoweredOn"								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Version)			= $ESX.Version + " - Build " + $ESX.Build	# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Date_Installation)	= "'" + $InstallDate						# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Serveurs_NTP)		= "$oESX_NTP"								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Zone_Temps)			= $UTC_Var1 + "+" + $UTC_Var2				# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Heure_Locale)		= $Esxcli2.Hardware.Clock.Get.Invoke()		# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Niveau_Compliance)	= $ComplianceLevel							# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Etat_Demon_NTP)		= $NTPD_status								# Ligne, Colonne
	
	### Colorisation des cellules
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Datacenter).Interior.ColorIndex 		= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Cluster).Interior.ColorIndex			= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Etat_ESX).Interior.ColorIndex			= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_vCenter).Interior.ColorIndex			= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Statut_ESX).Interior.ColorIndex			= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Date_Installation).Interior.ColorIndex	= $Excel_Couleur_Background
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Heure_Locale).Interior.ColorIndex		= $Excel_Couleur_Background
	
	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de la configuration SAN
Function Get-ESX_SAN { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
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
		$LUNs = Get-ScsiLun -Hba $hba  -LunType "disk" -ErrorAction SilentlyContinue
		$nrPaths = ($Target | %{$_.Lun.Count} | Measure-Object -Sum).Sum
		
		# Boucle déterminant les valeurs des variables pour chacune des interfaces HBA
		$ArraySAN_Line = (1..$Target_Count)
		ForEach($l in $ArraySAN_Line)	{
			$ArraySAN_HBAName	+= $hba.Name
			$ArraySAN_LUNs		+= $LUNS.Count
			$ArraySAN_Paths		+= $nrPaths
		}
		
		$Esxcli2 = Get-Esxcli -VMHost $ESX -Server $vCenter
		$ArraySAN_Paths_Active		+= ($Esxcli2.storage.core.path.list.invoke() | Where {($_.State -eq "active") -and ($_.Plugin -eq 'PowerPath') -and ($_.Adapter -eq $hba.Name)}).Count
		$ArraySAN_Paths_Active_INT 	+= ($Esxcli2.storage.core.path.list.invoke() | Where {($_.State -eq "active") -and ($_.Plugin -eq 'PowerPath') -and ($_.Adapter -eq $hba.Name)}).Count
		
		$ArraySAN_Paths_Dead		+= ($Esxcli2.storage.core.path.list.invoke() | Where {($_.State -eq "dead") -and ($_.Plugin -eq 'PowerPath') -and ($_.Adapter -eq $hba.Name)}).Count
		$ArraySAN_Paths_Dead_INT	+= ($Esxcli2.storage.core.path.list.invoke() | Where {($_.State -eq "dead") -and ($_.Plugin -eq 'PowerPath') -and ($_.Adapter -eq $hba.Name)}).Count
    }

	If ($ArraySAN_LUNs[1] -eq $ArraySAN_LUNs[2] -and $ArraySAN_Paths[1] -eq $ArraySAN_Paths[2]) { $PathsRedondance = "OK" }
	Else { $PathsRedondance = "NOK" }
	$PathsSum 			= ($ArraySAN_Paths[1] + $ArraySAN_Paths[2])
	$LUNsSum 			= ($ArraySAN_LUNs[1] + $ArraySAN_LUNs[2])
	$PathsDead_Total 	= ($ArraySAN_Paths_Dead_INT[0] + $ArraySAN_Paths_Dead_INT[1])
	$PathsActive_Total 	= ($ArraySAN_Paths_Active_INT[0] + $ArraySAN_Paths_Active_INT[1])
	
	$Global:PathsDead 	= "Total: $PathsDead_Total" + $vbcrlf + "(" + $ArraySAN_HBAName[1] + ": " + $ArraySAN_Paths_Dead[0] + ")" + $vbcrlf + "(" + $ArraySAN_HBAName[2] + ": " + $ArraySAN_Paths_Dead[1] + ")"
	$Global:PathsActive = "Total: $PathsActive_Total" + $vbcrlf + "(" + $ArraySAN_HBAName[1] + ": " + $ArraySAN_Paths_Active[0] + ")" + $vbcrlf + "(" + $ArraySAN_HBAName[2] + ": " + $ArraySAN_Paths_Active[1] + ")"
	
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_SAN_Total_Chemins) 			= "$PathsSum" + " (" + $Target_Count + " HBA)"		# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_SAN_Total_LUNs)				= "$LUNsSum" + " (" + $Target_Count + " HBA)"		# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_SAN_Total_Chemins_MORTS) 	= $PathsDead										# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_SAN_Total_Chemins_ACTIFS)	= $PathsActive										# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_SAN_Redondance_Chemins)		= $PathsRedondance									# Ligne, Colonne
	
	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de  configuration ESXi
Function Get-ESX_HOST { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
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

	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HOST_Current_CPU_Policy)		= $CPUPerformance											# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HOST_Mode_HA)				= $HAEnabled												# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HOST_Hyperthreading)			= $HyperThreading											# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HOST_EVC)					= "Current: " + $CurrentEVCMode #+ ", Max: " + $MaxEVCMode	# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HOST_Etat_Alarmes)			= $Alarm													# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HOST_TPS_Salting)			= $TPS														# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_HOST_Larges_Pages_RAM)		= $LPages													# Ligne, Colonne
	
	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de configuration du RESEAU
Function Get-ESX_NETWORK { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
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
	
	$Esxcli2 = Get-ESXCLI -VMHost $ESX -Server $vCenter
	$Global:ESXAdapter 		= $Esxcli2.Network.nic.list.Invoke()	| Where {$_.Link -eq "Up"}
	$Global:ESXAdapterName 	= $Esxcli2.Network.nic.list.Invoke()	| Where {$_.Link -eq "Up"} | Select Name, Speed, Duplex | Out-String
	$Global:ESXAdapterName 	= $ESXAdapterName -Replace "-","" -Replace "Name","" -Replace "Speed","" -Replace "Duplex","" -Replace "`r`n","" -Replace " ","" -Replace "10000", " 10000 " -Replace "vm", "`r`nvm"
	$Global:ESXAdapterCount = ($Esxcli2.Network.nic.list.Invoke() 	| Where {$_.Link -eq "Up"}).Count
	
	$vMotionIP = Get-VMHostNetworkAdapter -VMHost $ESX | Where {$_.DeviceName -eq "vmk1"} | Select IP | Out-String
	$vMotionIP = $vMotionIP -Replace "-","" -Replace "IP","" -Replace "`n","" -Replace " ",""
	$vMotionEnabled = (Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Summary.Config.VmotionEnabled
	If ($vMotionEnabled -eq $True) { $vMotionEnabled = "vMotion enabled" } Else { $vMotionEnabled = "vMotion disabled" }

	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_NETWORK_VLAN)		= $vLANID											# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_NETWORK_Adaptateurs) = "(" + $ESXAdapterCount + ") " + $ESXAdapterName	# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_NETWORK_vMotion)		= $vMotionIP + " (" + $vMotionEnabled + ")"			# Ligne, Colonne
	
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_NETWORK_vMotion).Interior.ColorIndex = $Excel_Couleur_Background
	
	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de configuration iLO
Function Get-HP_ILO { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "ILO`t`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs ILO")
	$ElapsedTime_Start = (Get-Date)
	
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ILO_Version) = $iLOversion # Ligne, Colonne
	
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

	$ArrayTag 		= @()	# Initialisation du tableau relatif au TAG
	$ArrayVersion 	= @()	# Initialisation du tableau relatif à la version ESXi
	$ArrayTimeZone	= @()	# Initialisation du tableau relatif à la TimeZone
	$ArrayBaseline	= @()	# Initialisation du tableau relatif à la Baseline
	$ArrayNTPd 		= @()	# Initialisation du tableau relatif au démon NTP
	$ArraySAN 		= @()	# Initialisation du tableau relatif au nombre de chemins SAN
	$ArrayLUN 		= @()	# Initialisation du tableau relatif au nombre de LUN
	$ArraySANDeath	= @()	# Initialisation du tableau relatif au nombre de chemins SAN morts
	$ArraySANAlive	= @()	# Initialisation du tableau relatif au nombre de chemins SAN alive
	$ArraySANRedon	= @()	# Initialisation du tableau relatif à la redondance SAN
	$ArrayCPUPolicy	= @()	# Initialisation du tableau relatif à l'utilisation CPU
	$ArrayHA		= @()	# Initialisation du tableau relatif au HA
	$ArrayHyperTh	= @()	# Initialisation du tableau relatif à l'HyperThreading
	$ArrayEVC		= @()	# Initialisation du tableau relatif au mode EVC
	$ArrayAlarme	= @()	# Initialisation du tableau relatif aux alarmes
	$ArrayTPS 		= @()	# Initialisation du tableau relatif au TPS
	$ArrayLPages	= @()	# Initialisation du tableau relatif aux Large Pages
	$ArrayVLAN 		= @()	# Initialisation du tableau relatif au nombre de VLAN ESX
	$ArrayLAN 		= @()	# Initialisation du tableau relatif au nombre d'adaptateurs LAN

	Get-ElapsedTime

	### Valorisation des tableaux relatifs aux colonnes à vérifier
	Write-Host " * Valorisation de la matrice `t`t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Valorisation de la matrice")
	$ElapsedTime_Start = (Get-Date)

	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
	ForEach($l in $ExcelLine)	{
		### Si le nom du cluster existe dans la feuille de référence ET que l'ESXi n'est pas arrêté
		If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Statut_ESX).Text -ne "PoweredOff") {
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_ESX).Text -eq "Connected") {
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Text -Like "@*") {
					$ArrayTag		+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Text
					$ArrayVersion	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Text
					$ArrayTimeZone 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Text
					$ArrayBaseline 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Text
					$ArrayNTPd	 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Text
					$ArraySAN 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.IndexOf(" "))
					$ArrayLUN		+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Text
					$ArraySANDeath	+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Text.IndexOf("(") - 1)
					$ArraySANAlive	+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Text.IndexOf("(") - 1)
					$ArraySANRedon	+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Text
					$ArrayCPUPolicy	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Text
					$ArrayHA		+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Text
					$ArrayHyperTh	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Text
					$ArrayEVC		+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Text
					$ArrayAlarme	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Text
					$ArrayTPS 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Text
					$ArrayLPages 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Text
					$ArrayVLAN 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1)
					$ArrayLAN 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Text.IndexOf(")") - 1)
				}
			}
		}
	}
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
	$Ref_ArraySANRedon	= ($ArraySANRedon	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayCPUPolicy	= ($ArrayCPUPolicy 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHA		= ($ArrayHA 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHyperTh	= ($ArrayHyperTh 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayEVC		= ($ArrayEVC 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayAlarme	= ($ArrayAlarme 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayTPS 		= ($ArrayTPS		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLPages	= ($ArrayLPages 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayVLAN 		= ($ArrayVLAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLAN 		= ($ArrayLAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	
	Get-ElapsedTime
	
	### Inscription des valeurs de références par cluster dans la 2ème feuille du fichier Excel
	Write-Host " * Inscription des valeurs de références dans Excel`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Inscription des valeurs de références dans Excel")
	$ElapsedTime_Start = (Get-Date)
	
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_Cluster)			= $Cluster.Name
	If ($Ref_ArrayVersion)		{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayVersion)		= $Ref_ArrayVersion		} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayVersion)	= "-" }
	If ($Ref_ArrayTimeZone) 	{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayTimeZone) 	= $Ref_ArrayTimeZone 	} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayTimeZone)	= "-" }
	If ($Ref_ArrayBaseline) 	{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayBaseline) 	= $Ref_ArrayBaseline 	} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayBaseline)	= "-" }
	If ($Ref_ArrayNTPd) 		{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayNTPd) 		= $Ref_ArrayNTPd 		} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayNTPd)		= "-" }
	If ($Ref_ArraySAN) 			{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArraySAN) 			= $Ref_ArraySAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArraySAN)		= "-" }
	If ($Ref_ArrayLUN) 			{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayLUN) 			= $Ref_ArrayLUN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayLUN)		= "-" }
	If ($Ref_ArrayCPUPolicy)	{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayCPUPolicy)	= $Ref_ArrayCPUPolicy 	} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayCPUPolicy)	= "-" }
	If ($Ref_ArrayHA) 			{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayHA) 			= $Ref_ArrayHA 			} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayHA)		= "-" }
	If ($Ref_ArrayHyperTh) 		{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayHyperTh) 		= $Ref_ArrayHyperTh 	} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayHyperTh)	= "-" }
	If ($Ref_ArrayEVC) 			{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayEVC) 			= $Ref_ArrayEVC 		} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayEVC)		= "-" }
	If ($Ref_ArrayAlarme) 		{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayAlarme) 		= $Ref_ArrayAlarme 		} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayAlarme)	= "-" }
	If ($Ref_ArrayTPS) 			{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayTPS) 			= $Ref_ArrayTPS 		} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayTPS)		= "-" }
	If ($Ref_ArrayLPages) 		{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayLPages)		= $Ref_ArrayLPages 		} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayLPages)	= "-" }
	If ($Ref_ArrayVLAN) 		{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayVLAN) 		= $Ref_ArrayVLAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayVLAN)		= "-" }
	If ($Ref_ArrayLAN) 			{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayLAN)			= $Ref_ArrayLAN 		} Else { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ArrayLAN)		= "-" }
	
	Switch -Wildcard ($vCenter)	{
		"*VCSZ*" { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ENVironnement) 	= "PRODUCTION" }
		"*VCSY*" { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ENVironnement) 	= "NON PRODUCTION" }
		"*VCSQ*" { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ENVironnement) 	= "CLOUD" }
		"*VCSZY*" { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ENVironnement) 	= "SITES ADMINISTRATIFS" }
		"*VCSSA*" { $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ENVironnement) 	= "BAC A SABLE" }
		Default	{ $ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1, $Excel_Ref_ENVironnement) 	= "NA" }
	}
	
	$ExcelWorkSheet.Cells.Item($ExcelWorkSheet_Ref.Cells.Item($Cluster_ID + 1), $Excel_Ref_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
	
	Get-ElapsedTime

	### Vérification des valeurs par rapport aux valeurs majoritaires
	Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Vérification des cellules vs valeurs majoritaires")
	$ElapsedTime_Start = (Get-Date)
	
	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ESX_Percent_NotCompliant_AVE = 0
	$ESX_NotCompliant_Item = 0
	
	$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
	ForEach($l in $ExcelLine)	{
		$ESX_Percent_NotCompliant = 0
			
		### Si le nom du cluster existe dans la feuille de référence ET que l'ESXi n'est pas arrêté
		If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Statut_ESX).Text -ne "PoweredOff") {
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_ESX).Text -eq "Connected") {
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Text -Like "@*") {
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Serveurs_NTP).Font.ColorIndex -eq $Excel_Couleur_Error)	{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[Erreur NTP Serveur]" + $vbcrlf; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Text -ne $Ref_ArrayTag) 							{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Tag/Ensemble]" + $vbcrlf; 				$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }								Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Font.ColorIndex = 1; 				$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Text -ne $Ref_ArrayVersion) 					{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Version/Ensemble]" + $vbcrlf; 			$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }							Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.ColorIndex = 1; 			$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Text -ne $Ref_ArrayTimeZone) 				{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ TimeZone/Ensemble]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }						Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.ColorIndex = 1; 			$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Text -ne $Ref_ArrayBaseline) 			{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Baseline/Ensemble]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }		Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = 1; 	$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Text -ne $Ref_ArrayNTPd) 				{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ NTP démon/Ensemble]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }				Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = 1; 		$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.IndexOf(" ")) -ne $Ref_ArraySAN) { $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= 					$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Chemins SAN/Ensemble]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Text -ne $Ref_ArrayLUN)						{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ LUNs/Ensemble]"; 			$ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 						Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.ColorIndex = 1; 			$ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Text.IndexOf("(") - 1) -ne $Ref_ArraySANDeath) { $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details) =$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[Chemin(s) mort(s)]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Text.IndexOf("(") - 1) -ne $Ref_ArraySANAlive) { $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details) =$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[Chemin(s) présent(s)]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Text -ne $Ref_ArraySANRedon) 			{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[Erreur redondance SAN]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }		Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Font.ColorIndex = 1; 	$ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Text -ne $Ref_ArrayCPUPolicy)		{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠CPU Policy/Ensemble]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }		Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = 1; 	$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Text -ne $Ref_ArrayHA) 							{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ HA/Ensemble]" + $vbcrlf; 				$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }							Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.ColorIndex = 1; 			$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Text -ne $Ref_ArrayHyperTh) 				{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ HyperThreading/Ensemble]" + $vbcrlf; 	$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }				Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.ColorIndex = 1; 		$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Text -ne $Ref_ArrayEVC) 							{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ EVC/Ensemble]" + $vbcrlf; 				$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 									Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.ColorIndex = 1; 				$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Text -ne $Ref_ArrayAlarme)					{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Alarmes/Ensemble]" + $vbcrlf; 			$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }					Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = 1; 		$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Text -ne $Ref_ArrayTPS) 					{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ TPS/Ensemble]" + $vbcrlf; 				$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }					Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.ColorIndex = 1; 		$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Text -ne $Ref_ArrayLPages) 			{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Large Pages/Ensemble]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }			Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = 1; 	$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1) -ne $Ref_ArrayVLAN) { $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= 						$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ VLANs/Ensemble]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.Bold = $False }
					If ($ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Text.IndexOf(")") - 1) -ne $Ref_ArrayLAN) { $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= 			$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Adaptateurs LAN/Ensemble]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Font.Bold = $False }
				
					### Inscription du pourcentage de conformité
					$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Globale)= "[" + [math]::round((100 - (($ESX_Percent_NotCompliant / 20) * 100)), 0) + "% conforme]"
					$ESX_Percent_NotCompliant_AVE = [math]::round(($ESX_Percent_NotCompliant_AVE + (100 - (($ESX_Percent_NotCompliant / 20) * 100))) / 2, 2)
				} Else {
					$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Globale) = "-"
					$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details) = "ESXi en standby..."	}
			} Else {
				$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Globale) = "-"
				$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details) = "ESXi en maintenance..."	}
		} Else {
			$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Globale) = "-"
			$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details) = "ESXi OFF..."	}
		
		### Mise à jour de la colonne TimeStamp
		$ExcelWorkSheet.Cells.Item($l, $Excel_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
	}
	
	Get-ElapsedTime
	
	Write-Host "Vérification de l'homogénéité des données du cluster... " -NoNewLine
	Write-Host "(Terminée)`r`n" -ForegroundColor Black -BackgroundColor White
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

	$ArrayTag 		= @()	# Initialisation du tableau relatif au TAG
	$ArrayVersion 	= @()	# Initialisation du tableau relatif à la version ESXi
	$ArrayTimeZone	= @()	# Initialisation du tableau relatif à la TimeZone
	$ArrayBaseline	= @()	# Initialisation du tableau relatif à la Baseline
	$ArrayNTPd 		= @()	# Initialisation du tableau relatif au démon NTP
	$ArraySAN 		= @()	# Initialisation du tableau relatif au nombre de chemins SAN
	$ArrayLUN 		= @()	# Initialisation du tableau relatif au nombre de LUNs
	$ArraySANDeath	= @()	# Initialisation du tableau relatif au nombre de chemins SAN morts
	$ArraySANAlive	= @()	# Initialisation du tableau relatif au nombre de chemins SAN alive
	$ArraySANRedon	= @()	# Initialisation du tableau relatif à la redondance SAN
	$ArrayCPUPolicy	= @()	# Initialisation du tableau relatif à l'utilisation CPU
	$ArrayHA		= @()	# Initialisation du tableau relatif au HA
	$ArrayHyperTh	= @()	# Initialisation du tableau relatif à l'HyperThreading
	$ArrayEVC		= @()	# Initialisation du tableau relatif au mode EVC
	$ArrayAlarme	= @()	# Initialisation du tableau relatif aux alarmes
	$ArrayTPS 		= @()	# Initialisation du tableau relatif au TPS
	$ArrayLPages	= @()	# Initialisation du tableau relatif aux Large Pages
	$ArrayVLAN 		= @()	# Initialisation du tableau relatif au nombre de VLAN ESX
	$ArrayLAN 		= @()	# Initialisation du tableau relatif au nombre d'adaptateurs LAN

	Get-ElapsedTime

	### Valorisation des tableaux relatifs aux colonnes à vérifier
	Write-Host " * Valorisation de la matrice `t`t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Valorisation de la matrice")
	$ElapsedTime_Start = (Get-Date)

	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
	ForEach($l in $ExcelLine)	{
		### Si le nom du cluster existe dans la feuille de référence ET que l'ESXi n'est pas arrêté
		If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Statut_ESX).Text -ne "PoweredOff") {
			$ArrayTag		+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Text
			$ArrayVersion	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Text
			$ArrayTimeZone 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Text
			$ArrayBaseline 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Text
			$ArrayNTPd	 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Text
			$ArraySAN 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.IndexOf(" "))
			$ArrayLUN 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Text
			$ArraySANDeath	+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Text.IndexOf("(") - 1)
			$ArraySANAlive	+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Text.IndexOf("(") - 1)
			$ArraySANRedon	+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Text
			$ArrayCPUPolicy	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Text
			$ArrayHA		+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Text
			$ArrayHyperTh	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Text
			$ArrayEVC		+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Text
			$ArrayAlarme	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Text
			$ArrayTPS 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Text
			$ArrayLPages 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Text
			$ArrayVLAN 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1)
			$ArrayLAN 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Text.IndexOf(")") - 1)
		}
	}

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
	$Ref_ArraySANRedon	= ($ArraySANRedon	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayCPUPolicy	= ($ArrayCPUPolicy 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHA		= ($ArrayHA 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayHyperTh	= ($ArrayHyperTh 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayEVC		= ($ArrayEVC 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayAlarme	= ($ArrayAlarme 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayTPS 		= ($ArrayTPS		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLPages	= ($ArrayLPages 	| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayVLAN 		= ($ArrayVLAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	$Ref_ArrayLAN 		= ($ArrayLAN 		| Group | ?{ $_.Count -gt ($ESX_Counter/2) }).Values | Select -Last 1
	
	Get-ElapsedTime
	
	### Vérification des valeurs par rapport aux valeurs majoritaires
	Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Vérification des cellules vs valeurs majoritaires")
	$ElapsedTime_Start = (Get-Date)
	
	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ESX_Percent_NotCompliant_AVE = 0
	$ESX_NotCompliant_Item = 0
	
	$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
	ForEach($l in $ExcelLine)	{
		$ESX_Percent_NotCompliant = 0
		
		### Si le nom du cluster existe dans la feuille de référence ET que l'ESXi n'est pas arrêté
		If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Statut_ESX).Text -ne "PoweredOff") {
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Serveurs_NTP).Font.ColorIndex -eq $Excel_Couleur_Error)	{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[Erreur NTP Serveur]" + $vbcrlf; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Text -ne $Ref_ArrayTag) 							{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Tag/Ensemble]" + $vbcrlf; 				$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 							Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Balise).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Text -ne $Ref_ArrayVersion) 					{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Version/Ensemble]" + $vbcrlf; 			$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }						Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Text -ne $Ref_ArrayTimeZone) 				{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ TimeZone/Ensemble]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }					Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Text -ne $Ref_ArrayBaseline) 			{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Baseline/Ensemble]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }	Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Text -ne $Ref_ArrayNTPd) 				{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ NTP démon/Ensemble]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }			Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.IndexOf(" ")) -ne $Ref_ArraySAN) { $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	=	$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Chemins SAN/Ensemble]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Text -ne $Ref_ArrayLUN) 						{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ LUNs/Ensemble]" + $vbcrlf;			 	$ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 					Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Text.IndexOf("(") - 1) -ne $Ref_ArraySANDeath) { $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details) =$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[Chemin(s) mort(s)]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_MORTS).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Text.IndexOf("(") - 1) -ne $Ref_ArraySANAlive) { $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details) =$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[Chemin(s) présent(s)]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins_ACTIFS).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Text -ne $Ref_ArraySANRedon) 			{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[Erreur redondance SAN]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 	Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Redondance_Chemins).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Text -ne $Ref_ArrayCPUPolicy)		{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ CPU Policy/Ensemble]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 	Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Text -ne $Ref_ArrayHA) 							{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ HA/Ensemble]" + $vbcrlf; 				$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 						Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Text -ne $Ref_ArrayHyperTh) 				{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ HyperThreading/Ensemble]" + $vbcrlf; 	$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 			Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Text -ne $Ref_ArrayEVC) 							{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ EVC/Ensemble]" + $vbcrlf; 				$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 								Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Text -ne $Ref_ArrayAlarme)					{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Alarmes/Ensemble]" + $vbcrlf; 			$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 				Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Text -ne $Ref_ArrayTPS) 					{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ TPS/Ensemble]" + $vbcrlf; 				$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 				Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Text -ne $Ref_ArrayLPages) 			{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Large Pages/Ensemble]" + $vbcrlf; 		$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 		Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1) -ne $Ref_ArrayVLAN) { $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= 	$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ VLANs/Ensemble]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.Bold = $False }
			If ($ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Text.IndexOf(")") - 1) -ne $Ref_ArrayLAN) { $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= 	$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Adaptateurs LAN/Ensemble]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } Else { $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Font.Bold = $False }
			
			### Inscription du pourcentage de conformité
			$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Globale)= "[" + [math]::round((100 - (($ESX_Percent_NotCompliant / 20) * 100)), 0) + "% conforme]"
			$ESX_Percent_NotCompliant_AVE = [math]::round(($ESX_Percent_NotCompliant_AVE + (100 - (($ESX_Percent_NotCompliant / 20) * 100))) / 2, 2)
		}
		Else {
			$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Globale) = "-"
			$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details) = "ESXi OFF..."
		}
		
		### Mise à jour de la colonne TimeStamp
		$ExcelWorkSheet.Cells.Item($l, $Excel_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
	}
	
	Get-ElapsedTime
	
	Write-Host "Vérification de l'homogénéité des données ESXi ciblés... " -NoNewLine
	Write-Host "(Terminée)`r`n" -ForegroundColor Black -BackgroundColor White
	Write-Host "--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Différence(s): $ESX_NotCompliant_Item`r`n * Taux moyen de conformité: $ESX_Percent_NotCompliant_AVE%`r`n------------------`r`n" -ForegroundColor White
	LogTrace ("Vérification de l'homogénéité des données du cluster... (Terminée)")
	Logtrace ("--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Différence(s): $ESX_NotCompliant_Item`r`n * Taux moyen de conformité: $ESX_Percent_NotCompliant_AVE%`r`n------------------`r`n")
}


### Fonction chargée de la comparaison/l'homogénéité des valeurs avec celles du fichier de référence
Function Get-ESX_Compare_CiblesvsReference {
	### Recherche du cluster relatif à l'ESXi ciblé dans la feuille de référence Excel
	$GetName = $ExcelWorkSheet_Ref.Range("A1").EntireColumn.Find("$Cluster")
	
	### Si le cluster n'existe pas dans la feuille de référence
	If (!($GetName.Row)) {
		Get-ESX_Compare_Cibles
		
		### Si les paramètres du fichier .BAT laissent supposer qu'il s'agit d'un nouveau cluster
		If ($args[4] -eq "TOUS" -and $args[2] -eq "AUCUN") {
			Write-Host "Les valeurs du fichier BAT correspondent à 'ESX_Inclus = TOUS' et 'ESX_Exclus = AUCUN'"
			LogTrace ("Les valeurs du fichier BAT correspondent à 'ESX_Inclus = TOUS' et 'ESX_Exclus = AUCUN'")
			Get-Ajout_Nouveau_Cluster
		}
		Else {
			Write-Host "Comparaison des ESXi ciblés impossible, cluster non trouvé... (Terminée)`r`n"
			Write-Host "--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Différence(s): $ESX_NotCompliant_Item`r`n------------------`r`n" -ForegroundColor White
			LogTrace ("Comparaison des ESXi ciblés impossible, cluster non trouvé... (Terminée)")
			LogTrace ("--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Différence(s): $ESX_NotCompliant_Item`r`n------------------`r`n")
		}
	}
	### Si le cluster existe dans la feuille de référence
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
		}
		
		Get-ElapsedTime

		### Vérification des données par rapport aux valeurs majoritaires
		Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
		LogTrace (" * Vérification des cellules vs valeurs majoritaires")
		$ElapsedTime_Start = (Get-Date)
			
		# Boucle déterminant la première ligne Excel du cluster à la dernière
		$ESX_Percent_NotCompliant_AVE = 0
		$ESX_NotCompliant_Item = 0

		$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
		ForEach($l in $ExcelLine)	{
			$ESX_Percent_NotCompliant = 0
			
			### Si le nom du cluster existe dans la feuille de référence ET que l'ESXi n'est pas arrêté
			If ($GetName.Row -and $ExcelWorkSheet.Cells.Item($GetName.Row, $Excel_CONF_Statut_ESX).Text -ne "PoweredOff") {
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Serveurs_NTP).Font.ColorIndex -eq $Excel_Couleur_Error)	{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[Erreur NTP Serveur]" + $vbcrlf; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Text -ne $Ref_ArrayVersion) 					{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Version/Référence]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }							Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Text -ne $Ref_ArrayTimeZone)					{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ TimeZone/Référence]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }						Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Text -ne $Ref_ArrayBaseline) 			{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Baseline/Référence]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }		Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Text -ne $Ref_ArrayNTPd) 				{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ NTP démon/Référence]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }			Else { $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.IndexOf(" ")) -ne $Ref_ArraySAN)		{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Chemins SAN/Référence]" + $vbcrlf; 	$ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }	Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Text -ne $Ref_ArrayLUN)						{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ LUNs/Référence]" + $vbcrlf; 	$ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }						Else { $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Text -ne $Ref_ArrayCPUPolicy)		{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ CPU Policy/Référence]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1}	Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Text -ne $Ref_ArrayHA) 							{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ HA/Référence]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 								Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Text -ne $Ref_ArrayHyperTh) 				{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ HyperThreading/Référence]" + $vbcrlf;$ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }		Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Text -ne $Ref_ArrayEVC) 							{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ EVC/Référence]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 } 										Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Text -ne $Ref_ArrayAlarme)					{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Alarmes/Référence]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }					Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Text -ne $Ref_ArrayTPS) 					{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ TPS/Référence]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }						Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Text -ne $Ref_ArrayLPages) 			{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ Large Pages/Référence]" + $vbcrlf; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1}		Else { $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Font.Bold = $False }
				If ($ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1) -ne $Ref_ArrayVLAN) 	{ $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details)	= $ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Details).Text + "[≠ VLAN/Référence]" + $vbcrlf; 			$ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.Bold = $True; $ESX_NotCompliant_Item += 1; $ESX_Percent_NotCompliant += 1 }	Else { $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Font.Bold = $False }

				### Inscription du pourcentage de conformité
				$ExcelWorkSheet.Cells.Item($l, $Excel_Conformite_Globale) = "[" + [math]::round((100 - (($ESX_Percent_NotCompliant / 15) * 100)), 0) + "% conforme]"
				$ESX_Percent_NotCompliant_AVE = [math]::round(($ESX_Percent_NotCompliant_AVE + (100 - (($ESX_Percent_NotCompliant / 15) * 100))) / 2, 2)
			}
			
			### Mise à jour de la colonne TimeStamp
			$ExcelWorkSheet.Cells.Item($l, $Excel_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
		}

		Get-ElapsedTime

		### Si le nom du cluster existe dans la feuille de référence
		If ($GetName.Row) {
			Write-Host "Comparaison des ESXi ciblés avec les références clusters... (Terminée)`r`n"
			Get-Ajout_Nouveau_Cluster
			Write-Host "--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Cluster trouvé à la ligne: $($GetName.Row)`r`n * Différence(s): $ESX_NotCompliant_Item`r`n * Taux moyen de conformité: $ESX_Percent_NotCompliant_AVE%`r`n------------------`r`n" -ForegroundColor White
			LogTrace ("Comparaison des ESXi ciblés avec les références clusters...... (Terminée)")
			LogTrace ("--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Cluster trouvé à la ligne: $($GetName.Row)`r`n * Différence(s): $ESX_NotCompliant_Item`r`n * Taux moyen de conformité: $ESX_Percent_NotCompliant_AVE%`r`n------------------`r`n")
		}
		### Si le nom du cluster n'existe pas dans la feuille de référence
		Else {
			### Si les paramètres du fichier .BAT laissent supposer qu'il s'agit d'un nouveau cluster
			If ($args[4] -eq "TOUS" -and $args[2] -eq "AUCUN") {
				Write-Host "Les valeurs du fichier BAT correspondent à 'ESX_Inclus = TOUS' et 'ESX_Exclus = AUCUN'"
				LogTrace ("Les valeurs du fichier BAT correspondent à 'ESX_Inclus = TOUS' et 'ESX_Exclus = AUCUN'")
				Get-Ajout_Nouveau_Cluster
			}
			Else {
				Write-Host "Comparaison des ESXi ciblés impossible, cluster non trouvé... (Terminée)`r`n"
				Write-Host "--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Différence(s): $ESX_NotCompliant_Item`r`n------------------`r`n" -ForegroundColor White
				LogTrace ("Comparaison des ESXi ciblés impossible, cluster non trouvé... (Terminée)")
				LogTrace ("--- En résumé: ---`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Différence(s): $ESX_NotCompliant_Item`r`n------------------`r`n")
			}
		}
	}
}


### Fonction chargée de la comparaison/l'homogénéité des valeurs entre elles
Function Get-Ajout_Nouveau_Cluster {
	Write-Host "Ajout/Mise à jour du cluster dans les références... " -NoNewLine
	Write-Host "En cours" -ForegroundColor Green
	LogTrace ("Ajout/Mise à jour du cluster dans les références...")

	### Initialisation de la matrice 2 dimensions
	Write-Host " * Initialisation de la matrice `t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Initialisation de la matrice")
	$ElapsedTime_Start = (Get-Date)

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

	Get-ElapsedTime

	### Valorisation des tableaux relatifs aux colonnes à vérifier
	Write-Host " * Valorisation de la matrice `t`t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Valorisation de la matrice")
	$ElapsedTime_Start = (Get-Date)
	
	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = ($ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1))
	ForEach($l in $ExcelLine)	{
		$ArrayVersion	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Version).Text
		$ArrayTimeZone 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Zone_Temps).Text
		$ArrayBaseline 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Niveau_Compliance).Text
		$ArrayNTPd	 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_CONF_Etat_Demon_NTP).Text
		$ArraySAN 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_Chemins).Text.IndexOf(" "))
		$ArrayLUN 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_SAN_Total_LUNs).Text
		$ArrayCPUPolicy	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Current_CPU_Policy).Text
		$ArrayHA		+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Mode_HA).Text
		$ArrayHyperTh	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Hyperthreading).Text
		$ArrayEVC		+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_EVC).Text
		$ArrayAlarme	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Etat_Alarmes).Text
		$ArrayTPS 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_TPS_Salting).Text
		$ArrayLPages 	+= $ExcelWorkSheet.Cells.Item($l, $Excel_HOST_Larges_Pages_RAM).Text
		$ArrayVLAN 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_VLAN).Text.IndexOf(")") - 1)
		$ArrayLAN 		+= $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, $Excel_NETWORK_Adaptateurs).Text.IndexOf(")") - 1)
	}
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
	
	Get-ElapsedTime
	
	### Inscription des valeurs de références par cluster dans la 2ème feuille du fichier Excel
	Write-Host " * Inscription des valeurs de références dans Excel`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Inscription des valeurs de références dans Excel")
	$ElapsedTime_Start = (Get-Date)
	
	### Si les paramètres du fichier .BAT laissent supposer qu'il s'agit d'un nouveau cluster
	If ($args[4] -eq "TOUS" -and $args[2] -eq "AUCUN") {
		### Recherche d'une cellule vide dans la feuille de référence Excel
		$GetName = $ExcelWorkSheet_Ref.Range("A1").EntireColumn.Find($Cluster.Name)
		
		Write-Host "Mise à jour des valeurs de références du cluster à la ligne $GetName.Row"
		LogTrace ("Mise à jour des valeurs de références du cluster à la ligne $GetName.Row")
		
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
	}
	 
	Switch -Wildcard ($vCenter)	{
		"*VCSZ*" { $ExcelWorkSheet.Cells.Item($GetName.Row, $Excel_Ref_ENVironnement)	= "PRODUCTION" }
		"*VCSY*" { $ExcelWorkSheet.Cells.Item($GetName.Row, $Excel_Ref_ENVironnement) 	= "NON PRODUCTION" }
		"*VCSQ*" { $ExcelWorkSheet.Cells.Item($GetName.Row, $Excel_Ref_ENVironnement) 	= "CLOUD" }
		"*VCSZY*" { $ExcelWorkSheet.Cells.Item($GetName.Row, $Excel_Ref_ENVironnement) 	= "SITES ADMINISTRATIFS" }
		"*VCSSA*" { $ExcelWorkSheet.Cells.Item($GetName.Row, $Excel_Ref_ENVironnement) 	= "BAC A SABLE" }
		Default	{ $ExcelWorkSheet.Cells.Item($GetName.Row, $Excel_Ref_ENVironnement) 	= "NA" }
	}

	### Mise à jour de la colonne TimeStamp
	$ExcelWorkSheet_Ref.Cells.Item($GetName.Row, $Excel_Ref_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
	
	Get-ElapsedTime
}


### Fonction chargée de la comparaison/l'homogénéité des clusters (En fin de traitement complet)
Function Get-Cluster_Compare {
	Write-Host "Vérification de l'homogénéité des clusters par ENV... " -NoNewLine
	Write-Host "En cours" -ForegroundColor Green
	LogTrace ("Vérification de l'homogénéité des clusters par ENV")
	
	### Récupération des valeurs majoritaires par cluster et par ENV
	Write-Host " * Initialisation de la matrice `t`t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Initialisation de la matrice")
	$ElapsedTime_Start = (Get-Date)
	
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
	
	Get-ElapsedTime
	
	### Valorisation des tableaux relatifs aux colonnes à vérifier
	Write-Host " * Valorisation de la matrice`t`t`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Valorisation de la matrice")
	$ElapsedTime_Start = (Get-Date)
	
	### Recherche d'une cellule vide dans la feuille de référence Excel
	$GetName = $ExcelWorkSheet_Ref.Range("A1").EntireColumn.Find("")

	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = (2..($GetName.Row))
	ForEach($l in $ExcelLine)	{
		$ArrayVersion	+= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayVersion).Text
		$ArrayTimeZone 	+= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTimeZone).Text
		$ArrayBaseline 	+= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayBaseline).Text
		$ArrayNTPd	 	+= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayNTPd).Text
		$ArrayCPUPolicy	+= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayCPUPolicy).Text
		$ArrayHA		+= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHA).Text
		$ArrayHyperTh	+= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHyperTh).Text
		$ArrayEVC		+= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayEVC).Text
		$ArrayAlarme	+= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayAlarme).Text
		$ArrayTPS 		+= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTPS).Text
		$ArrayLPages 	+= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayLPages).Text
	}
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
	
	Get-ElapsedTime
	
	### Vérification des données par rapport aux valeurs majoritaires
	Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Vérification des cellules vs valeurs majoritaires")
	$ElapsedTime_Start = (Get-Date)
	
	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = (2..($GetName.Row))
	ForEach($l in $ExcelLine)	{
		$Global:Cluster_NotCompliant_Item = 0
		
		If ($ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayVersion).Text -ne $Ref_ArrayVersion) 	{ $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details).Text + "[≠ Version/Ensemble]" + $vbcrlf; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayVersion).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayVersion).Font.Bold = $True; $Cluster_NotCompliant_Item += 1 } 		Else { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayVersion).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayVersion).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTimeZone).Text -ne $Ref_ArrayTimeZone) 	{ $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details).Text + "[≠ TimeZone/Ensemble]" + $vbcrlf; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTimeZone).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTimeZone).Font.Bold = $True; $Cluster_NotCompliant_Item += 1 } 		Else { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTimeZone).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTimeZone).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayBaseline).Text -ne $Ref_ArrayBaseline) 	{ $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details).Text + "[≠ Baseline/Ensemble]" + $vbcrlf; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayBaseline).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayBaseline).Font.Bold = $True; $Cluster_NotCompliant_Item += 1 } 		Else { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayBaseline).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayBaseline).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayNTPd).Text -ne $Ref_ArrayNTPd) 			{ $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details).Text + "[≠ NTP démon/Ensemble]" + $vbcrlf; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayNTPd).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayNTPd).Font.Bold = $True; $Cluster_NotCompliant_Item += 1 } 			Else { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayNTPd).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayNTPd).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayCPUPolicy).Text -ne $Ref_ArrayCPUPolicy) { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details).Text + "[≠ CPU Policy/Ensemble]" + $vbcrlf; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayCPUPolicy).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayCPUPolicy).Font.Bold = $True; $Cluster_NotCompliant_Item += 1 }	Else { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayCPUPolicy).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayCPUPolicy).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHA).Text -ne $Ref_ArrayHA) 				{ $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details).Text + "[≠ HA/Ensemble]" + $vbcrlf; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHA).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHA).Font.Bold = $True; $Cluster_NotCompliant_Item += 1 } 						Else { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHA).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHA).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHyperTh).Text -ne $Ref_ArrayHyperTh) 	{ $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details).Text + "[≠ HyperThreading/Ensemble]" + $vbcrlf; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHyperTh).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHyperTh).Font.Bold = $True; $Cluster_NotCompliant_Item += 1 } 	Else { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHyperTh).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayHyperTh).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayEVC).Text -ne $Ref_ArrayEVC) 			{ $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details).Text + "[≠ EVC/Ensemble]" + $vbcrlf; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayEVC).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayEVC).Font.Bold = $True; $Cluster_NotCompliant_Item += 1 } 					Else { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayEVC).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayEVC).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayAlarme).Text -ne $Ref_ArrayAlarme)		{ $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details).Text + "[≠ Alarmes/Ensemble]" + $vbcrlf; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayAlarme).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayAlarme).Font.Bold = $True; $Cluster_NotCompliant_Item += 1 } 			Else { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayAlarme).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayAlarme).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTPS).Text -ne $Ref_ArrayTPS) 			{ $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details).Text + "[≠ TPS/Ensemble]" + $vbcrlf; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTPS).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTPS).Font.Bold = $True; $Cluster_NotCompliant_Item += 1 } 					Else { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTPS).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayTPS).Font.Bold = $False }
		If ($ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayLPages).Text -ne $Ref_ArrayLPages) 		{ $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details)	= $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_Conformite_Details).Text + "[≠ Large Pages/Ensemble]" + $vbcrlf; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayLPages).Font.ColorIndex = $Excel_Couleur_Error; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayLPages).Font.Bold = $True; $Cluster_NotCompliant_Item += 1 } 		Else { $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayLPages).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_ArrayLPages).Font.Bold = $False }
	
		### Mise à jour de la colonne TimeStamp
		$ExcelWorkSheet_Ref.Cells.Item($l, $Excel_Ref_TimeStamp) = Get-Date -UFormat "%Y/%m/%d %H:%M:%S"
	}
	
	Get-ElapsedTime
	
	Write-Host "Vérification de l'homogénéité des clusters par ENV... (Terminée)`r`n"
	Write-Host "En résumé: $Cluster_NotCompliant_Item différence(s)`r`n" -ForegroundColor White
	LogTrace ("Vérification de l'homogénéité des clusters par ENV... (Terminée)")
	Logtrace ("En résumé: $Cluster_NotCompliant_Item différence(s)`r`n")
}


### Fonction chargée de l'envoi de messages électroniques (En fin de traitement)
Function Send_Mail {
	$Sender1 = "Christophe.VIDEAU-ext@ca-ts.fr"
	$Sender2 = #"MCO.Infra.OS.distribues@ca-ts.fr"
	$From = "[ESXi] compliance check <ESXiCompliance.report@ca-ts.fr>"
	$Subject = "[Conformit&eacute; ESXi] Compte-rendu operationnel {Conformit&eacute; infrastructures VMware}"
	If ($BodyMail_Error -ne $NULL) { $Body = "Bonjour,<BR><BR>Vous trouverez en pi&egrave;ce jointe le fichier LOG et la synth&egrave;se technique de l&rsquo;op&eacute;ration.<BR><BR><U>En bref:</U><BR>Mode: " + $Mode + "<BR>Nombre d'ESXi conformes:<B><I> " + $ESX_Compliant + "</I></B><BR>Nombre d'ESXi non conforme(s):<B><I> " + $ESX_NotCompliant + " (" + $ESX_NotCompliant_Item + " diff&eacute;rences depuis le début)</I><BR>Nombre de clusters non conforme(s):<B><I> " + $Cluster_NotCompliant + " (" + $Cluster_NotCompliant_Item + " diff&eacute;rences)</I></B><BR>Dur&eacute;e du traitement: <I>" + $($Min) + "min. " + $($Sec) + "sec.</I><BR><BR>----------------------" + $BodyMail_Error + "<BR>----------------------<BR><BR><BR>Cordialement.<BR>L'&eacute;quipe d&eacute;veloppement (Contact: Hugues de TERSSAC)<BR><BR>Ce message a &eacute;t&eacute; envoy&eacute; par un automate, ne pas y r&eacute;pondre." }
	Else { $Body = "Bonjour,<BR><BR>Vous trouverez en pi&egrave;ce jointe le fichier LOG et la synth&egrave;se technique de l&rsquo;op&eacute;ration.<BR><BR><U>En bref:</U><BR>Mode: " + $Mode + "<BR>Nombre d'ESXi conformes:<B><I> " + $ESX_Compliant + "</I></B><BR>Nombre d'ESXi non conforme(s):<B><I> " + $ESX_NotCompliant + " (" + $ESX_NotCompliant_Item + " diff&eacute;rences depuis le début)</I></B><BR>Nombre de clusters non conforme(s):<B><I> " + $Cluster_NotCompliant + " (" + $Cluster_NotCompliant_Item + " diff&eacute;rences)</I><BR>Dur&eacute;e du traitement: <I>" + $($Min) + "min. " + $($Sec) + "sec.</I><BR><BR>Cordialement.<BR>L'&eacute;quipe d&eacute;veloppement (Contact: Christophe VIDEAU)<BR><BR>Ce message a &eacute;t&eacute; envoy&eacute; par un automate, ne pas y r&eacute;pondre." }
	$Attachments = $FicLog, $FicRes
	$SMTP = "muz10-e1smtp-IN-DC-INT.zres.ztech"
	Send-MailMessage -To $Sender1, $Sender2 -From $From -Subject $Subject -Body $Body -Attachments $Attachments -SmtpServer $SMTP -Priority High -BodyAsHTML
}


### >>>>>>>>>> SOUS-FONCTIONS <<<<<<<<<<<<
Function LogTrace ($Message){
	$Message = (Get-Date -format G) + " " + $Message
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Message -Append }

### >>>>>>> DEBUT DU SCRIPT <<<<<<<<<<<
#$ErrorActionPreference = "SilentlyContinue"
$ScriptName	= [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition)
$vbcrlf		= "`r`n"
$l			= 0
$dat		= Get-Date -UFormat "%Y_%m_%d_%H_%M_%S"

### Création du répertoire de LOG si besoin
$PathScript = ($myinvocation.MyCommand.Definition | Split-Path -parent) + "\"
$RepLog     = $PathScript + "LOG"
If (!(Test-Path $RepLog)) { New-Item -Path $PathScript -ItemType Directory -Name "LOG"}
$RepLog     = $RepLog + "\"
If ($args[1]) { $FicLog     = $RepLog + $ScriptName + "_" + $dat + "_ONE_SHOT.log" } Else { $FicLog     = $RepLog + $ScriptName + "_REF.log" }
$LineSep    = "=" * 70

### Si le fichier LOG n'existe pas on le crée à vide
$Line = ">> DEBUT script de contrôle ESXi <<"
If (!(Test-Path $FicLog)) {
	$Line = (Get-Date -format G) + " " + $Line
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Line }
Else { LogTrace ($Line) }

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
	
	$FicRes		= $RepLog + $ScriptName + "_" + $dat + "_ONE_SHOT.xlsx"
	$FicRes_Ref	= $RepLog + $ScriptName + "_REF.xlsx"
	
	If (Test-Path ($RepLog + $ScriptName + "_REF.xlsx")) {
		Write-Host "Le fichier Excel de référence '$($RepLog + $ScriptName + "_REF.xlsx")' est disponible" -ForegroundColor Green
		Write-Host "Comparaison possible en fin de traitement de chacun des clusters..." -ForegroundColor White
		LogTrace ("Le fichier Excel de référence '$($RepLog + $ScriptName + "_REF.xlsx")' est disponible.`r`nComparaison possible en fin de traitement de chacun des clusters...")
		$ExcelWorkBook_Ref	= $Excel.WorkBooks.Open($FicRes_Ref)	# Ouverture du fichier *_REF.xlsx
		$ExcelWorkSheet_Ref	= $Excel.WorkSheets.item(2)				# Définition de la feuille Excel par défaut du fichier *_REF.xlsx
	} Else {
		Write-Host "INFO: Le fichier Excel de référence '$RepLog + $ScriptName + "_REF.xlsx"' est indisponible.`r`nComparaison impossible en fin de traitement de chacun des clusters..." -ForegroundColor Red
		LogTrace ("INFO: Le fichier Excel de référence '$($RepLog + $ScriptName + "_REF.xlsx")' est indisponible.`r`nComparaison impossible en fin de traitement de chacun des clusters...") }
	
	Copy-Item ($PathScript + "ESX_Health_Modele.xlsx") -Destination ($RepLog + $ScriptName + "_" + $dat + "_ONE_SHOT.xlsx")
	
	### Définition du fichier Excel
	$ExcelWorkBook		= $Excel.WorkBooks.Open($FicRes)		# Ouverture du fichier *_[ONE SHOT].xlsx
	$ExcelWorkSheet		= $Excel.WorkSheets.item(1)				# Définition de la feuille Excel par défaut du fichier *_[ONE SHOT].xlsx
	$Excel.WorkSheets.item(2).Delete()							# Suppression de la 2ème feuille (inutile) du fichier *_[ONE SHOT].xlsx
	
	$ExcelWorkSheet.Activate()
	$Excel.Visible		= $False
	
	$Global:vCenters	= $args[0] # Nom du vCenter
	$Cluster_Exc 		= $args[1] # clusters exclusq
	$ESX_Exc			= $args[2] # ESXi exclus
	$Cluster_Inc		= $args[3] # clusters inclus
	$ESX_Inc			= $args[4] # ESXi inclus

	### Bouchon pour tests
	<# $vcenters	= "swmuzv1vcszd.zres.ztech"
	$Cluster_Exc	= "AUCUN"
	$ESX_Exc		= "AUCUN"
	$Cluster_Inc	= "CL_MU_HDI_Z80,CL_MU_HDM_Z80"
	$ESX_Inc		= "sxmuzhvhdich.zres.ztech,sxmuzhvhdidk.zres.ztech" #>

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

	LogTrace ("Mode 'Par vCenter avec filtrage'")
	LogTrace ("Fichier liste des vCenter à traiter .. : $vCenters")
	Logtrace ("Cluster à exclure .................... : $Cluster_Exc")
	Logtrace ("ESXi à exclure ....................... : $ESX_Exc")
	Logtrace ("Cluster à prendre en compte .......... : $Cluster_Inc")
	Logtrace ("ESXi à prendre en compte ............. : $ESX_Inc")
	LogTrace ($LineSep + $vbcrlf)
	$TabVcc = $vCenters.split(",")
} Else {
	$Mode = "Mode 'Par vCenter sans aucun filtrage'"
	$Mode_Var = 1
	Write-Host "$Mode [ID $Mode_Var]" -ForegroundColor Red -BackgroundColor White
	
	### Test de l'existence du fichier Excel de référence
	If (Test-Path ($RepLog + $ScriptName + "_REF.xlsx")) {
		$Reponse = [System.Windows.Forms.MessageBox]::Show("Attention: le fichier '$($RepLog + $ScriptName + "_REF.xlsx")' existe déjà, voulez-vous le remplacer ? ", "Confirmation" , 4, 32)
		If (($Reponse -eq "Yes") -or ($Reponse -eq "Oui")) {
			Move-Item ($RepLog + $ScriptName + "_REF.xlsx") ($RepLog + $ScriptName + "_REF_(BCKP_"+ $dat + ").xlsx")	# Sauvegarde du précédent fichier _REF.xlsx
			Move-Item ($RepLog + $ScriptName + "_REF.log") ($RepLog + $ScriptName + "_REF_(BCKP_"+ $dat + ").log")		# Sauvegarde du précédent fichier _REF.log
			Copy-Item ($PathScript + "ESX_Health_Modele.xlsx") -Destination ($RepLog + $ScriptName + "_REF.xlsx")		# Copie du modèle pour modification par le script
			Write-Host "Le fichier '$($RepLog + $ScriptName + "_REF.xlsx")' (et son fichier LOG) a été sauvegardé" -ForegroundColor Red
			LogTrace ("Le fichier '$($RepLog + $ScriptName + "_REF.xlsx")' (et son fichier LOG) a été sauvegardé") }
		Else	{
			Write-Host "Le fichier '$($RepLog + $ScriptName + "_REF.xlsx")' n'a pas été écrasé" -ForegroundColor Red
			LogTrace ("Le fichier '$($RepLog + $ScriptName + "_REF.xlsx")' n'a pas été écrasé")
			LogTrace ("FIN du script")
			Write-Host -NoNewLine "FIN du script. Appuyer sur une touche pour quitter...`r`n"
			$NULL = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
			Exit }
	} Else { Copy-Item ($PathScript + "ESX_Health_Modele.xlsx") -Destination ($RepLog + $ScriptName + "_REF-work.xlsx") }	# Copie du modèle pour modification par le script
	
	$FicRes = $RepLog + $ScriptName + "_REF-work.xlsx"
	
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
	LogTrace ($LineSep + $vbcrlf)
	$TabVcc = $vCenters.split(",") }

$user       = "CSP_SCRIPT_ADM"
$fickey     = "D:\Scripts\Credentials\key.crd"
$ficcred    = "D:\Scripts\Credentials\vmware_adm.crd"
$key        = Get-content $fickey
$pwd        = Get-Content $ficcred | ConvertTo-SecureString -key $key
$Credential = New-Object System.Management.Automation.PSCredential $user, $pwd


### Boucle de traitement des vCenters contenus dans les paramètres de la ligne de commandes
ForEach ($vCenter in $TabVcc) {
	LogTrace ("DEBUT du traitement VCENTER $vCenter")
	Write-Host "`r`nDEBUT du traitement VCENTER " -NoNewLine
	Write-Host "$vCenter... ".ToUpper() -ForegroundColor Yellow -NoNewLine
	Write-Host "En cours" -ForegroundColor Green	
	
	$rccnx = Connect-VIServer -Server $vCenter -Protocol https -Credential $Credential
	$topCnxVcc = "0"
	If ($rccnx -ne $NULL) { If ($rccnx.Isconnected) { $topCnxVcc = "1" } }

	If ($topCnxVcc -ne "1") { LogTrace ("ERREUR: Connexion KO au vCenter $vCenter - Arrêt du script")
		Write-Host "ERREUR: Connexion KO au vCenter $vCenter - Arrêt du script" -ForegroundColor White -BackgroundColor Red	
		$rc += 1
		Exit $rc }
	Else { LogTrace ("SUCCES: Connexion OK au vCenter $vCenter" + $vbcrlf)
		Write-Host "SUCCES: Connexion OK au vCenter $vCenter" -ForegroundColor Black -BackgroundColor Green }
	
	$Global:noDatacenter = 0
	$Global:oDatacenters = Get-Datacenter | Sort Name
	$Global:Datacenter_Counter = $oDatacenters.Count
	
	### Boucle de traitement des Datacenter composant le vCenter
	ForEach($DC in $oDatacenters){ $noDatacenter += 1
		LogTrace ("Traitement DATACENTER $DC n°$noDatacenter sur $Datacenter_Counter" + $vbcrlf)
		Write-Host "`r`nTraitement DATACENTER [#$noDatacenter/$Datacenter_Counter] " -NoNewLine
		Write-Host "$DC... ".ToUpper() -ForegroundColor Yellow -NoNewLine
		Write-Host "En cours" -ForegroundColor Green	

		$Global:noCluster = 0
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

			$noCluster += 1; $Cluster_ID += 1
			### Exception de valorisation de la variable selon le mode
			If ($Mode_Var -eq 1) {
				$Global:ExcelLine_Start += $ESX_Counter
			}
			Else {
				$Global:ExcelLine_Start += $no_inc_ESX
			}

			LogTrace ("Traitement CLUSTER '$Cluster' n°$noCluster sur $Cluster_Counter {$Cluster_ID}")
			Write-Host "`r`nTraitement CLUSTER [#$noCluster/$Cluster_Counter] " -NoNewLine
			Write-Host "{$Cluster_ID} " -ForegroundColor Red -NoNewLine
			Write-Host "'$Cluster'... ".ToUpper() -ForegroundColor Yellow -NoNewLine
			Write-Host "En cours" -ForegroundColor Green

			$Global:oESX = Get-vmHost -Location $Cluster
			$Global:ESX_Counter = $oESX.Count
			$Global:no_ESX = 0

			### Boucle de traitement des ESXi composants le cluster dont ceux contenus dans les paramètres de la ligne de commandes
			ForEach($ESX in Get-vmHost -Location $Cluster) {
				If (($Mode_Var -eq 2) -and ($ESX_Inc -ne "TOUS")) {
					$Global:ESX_Counter = $Array_ESX_Inc_Counter
					If ($Array_ESX_Exc -Contains $ESX) {
						Logtrace (" * Exclusion de l'ESXi '$ESX'...")
						Write-Host " * Exclusion " -NoNewLine -ForegroundColor Red
						Write-Host "de l'ESXi " -NoNewLine
						Write-Host "'$ESX'..." -ForegroundColor DarkYellow
						Continue
					}
					If (($Array_ESX_Inc.Length -ne 0) -and ($Array_ESX_Inc -notContains $ESX)) {
						Logtrace (" * ESXi '$ESX' absent des ESXi à traiter...")
						Write-Host " * ESXi " -NoNewLine
						Write-Host "'$ESX' " -NoNewLine -ForegroundColor DarkYellow
						Write-Host "absent " -NoNewLine -ForegroundColor Red
						Write-Host "des ESXi à traiter..."
						Continue
					}
				}
 
				$no_ESX += 1; $no_inc_ESX += 1

				$StartTime = Get-Date -Format HH:mm:ss
				LogTrace ("Traitement ESXi '$ESX' n°$no_ESX sur $($ESX_Counter) {$no_inc_ESX}")
				Write-Host "[$StartTime] Traitement ESXi [#$no_ESX/$($ESX_Counter)] " -NoNewLine
				Write-Host "{$no_inc_ESX} " -ForegroundColor Red -NoNewLine
				Write-Host "'$ESX'... ".ToUpper() -ForegroundColor Yellow -NoNewLine
				Write-Host "En cours" -ForegroundColor Green

				If ($ESX.PowerState -ne "PoweredOn") {
					For ($i = 1; $i -le $Excel_Conformite_Details; $i++) {
						$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $i) = "NA"
						$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $i).Interior.ColorIndex = $Excel_Couleur_Background
					}
					$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_Nom_ESX)			= $ESX.Name			# Ligne, Colonne
					$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Statut_ESX) = "PoweredOff" 		# Ligne, Colonne
					$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Datacenter)	= "$DC"				# Ligne, Colonne
					$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_Cluster)	= "$Cluster"		# Ligne, Colonne
					$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_CONF_vCenter)	= "$vCenter"		# Ligne, Colonne

					Switch -Wildcard ($vCenter)	{
						"*VCSZ*" { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement)	= "PRODUCTION" }
						"*VCSY*" { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement)	= "NON PRODUCTION" }
						"*VCSQ*" { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement)	= "CLOUD" }
						"*VCSZY*" { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement) 	= "SITES ADMINISTRATIFS" }
						"*VCSSA*" { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement)	= "BAC A SABLE" }
						Default	{ $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, $Excel_ENVironnement)		= "NA" }
					}

					$EndTime = Get-Date -Format HH:mm:ss
					
					### Enregistrement des modifications Excel
					$ExcelWorkBook.Save()
					
					LogTrace ("MISE A JOUR des données pour l'ESXi '$ESX'. Zone [L$ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1) *L$($no_inc_ESX + 1)*]..." + $vbcrlf)
					Write-Host "[$EndTime] Mise à jour des données Excel " -NoNewLine
					Write-Host "Zone [L$ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1) *L$($no_inc_ESX + 1)*]... "  -ForegroundColor Yellow -NoNewLine
					Write-Host "Terminée`r`n" -ForegroundColor Black -BackgroundColor White

					# Dans le cas d'une découverte complète du périmètre
					If ($Mode_Var -eq 1) {
						If ($no_ESX -eq $($oESX.Count)) {
							Get-ESX_Compare_Full; $ExcelWorkBook.Save()
						}
					}
					
					### Sélection de la fonction selon le mode de départ
					If ($Mode_Var -eq 2) {
						If (Test-Path ($RepLog + $ScriptName + "_REF.xlsx")) {	# Dans le cas d'une vérification de conformité unitaire si le fichier Excel de référence existe
							If ($no_ESX -eq $ESX_Counter) {
								Get-ESX_Compare_CiblesvsReference
								$ExcelWorkBook.Save()
								Break
							}
						}
						Else {	# Dans le cas d'une vérification de conformité unitaire si le fichier Excel de référence n'existe pas
							If ($no_ESX -eq $ESX_Counter) {
								Get-ESX_Compare_Cibles
								$ExcelWorkBook.Save()
								Break
							}
						}
					}
					Continue
				}
				
				### Exécution des fonctions de récupération des données ESX
				Get-ESX_HARD -vmHost $ESX		# Récupération matérielle
				Get-ESX_CONFIG -vmHost $ESX		# Récupération de la configuration matérielle
				Get-ESX_SAN -vmHost $ESX		# Récupération des données stockage SAN
				Get-ESX_HOST -vmHost $ESX		# Récupération de la configuration ESXi
				Get-ESX_NETWORK -vmHost $ESX	# Récupération de la configuration réseaux
				Get-HP_ILO -vmHost $ESX			# Récupération de la configuration iLO
				
				### Enregistrement des modifications Excel
				$ExcelWorkBook.Save()

				$EndTime = Get-Date -Format HH:mm:ss
				LogTrace ("MISE A JOUR des données pour l'ESXi '$ESX'. Zone [L$ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1) *L$($no_inc_ESX + 1)*]..." + $vbcrlf)
				Write-Host "[$EndTime] Mise à jour des données Excel " -NoNewLine
				Write-Host "Zone [L$ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1) *L$($no_inc_ESX + 1)*]... "  -ForegroundColor Yellow -NoNewLine
				Write-Host "Terminée`r`n" -ForegroundColor Black -BackgroundColor White

				# Dans le cas d'une découverte complète du périmètre
				If ($Mode_Var -eq 1) {
					If ($no_ESX -eq $($oESX.Count)) {
						Get-ESX_Compare_Full; $ExcelWorkBook.Save()
					}
				}
				
				### Sélection de la fonction selon le mode de départ
				If ($Mode_Var -eq 2) {
					If (Test-Path ($RepLog + $ScriptName + "_REF.xlsx")) {	# Dans le cas d'une vérification de conformité unitaire si le fichier Excel de référence existe
						If ($no_ESX -eq $ESX_Counter) {
							Get-ESX_Compare_CiblesvsReference
							$ExcelWorkBook.Save()
							Break
						}
					}
					Else {	# Dans le cas d'une vérification de conformité unitaire si le fichier Excel de référence n'existe pas
						If ($no_ESX -eq $ESX_Counter) {
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
	If ($Mode_Var -eq 1) { Get-Cluster_Compare }
	
	LogTrace ("DECONNEXION et FIN du traitement depuis le VCENTER '$vCenter'`r`n")
	Disconnect-VIServer -Server $vCenter -Force -Confirm:$False
}

$ExcelWorkBook.Save()
Write-Host "Fermeture du classeur Excel [Terminé]"
LogTrace ("Fermeture du classeur Excel [Terminé]")

$Excel.Quit()
Write-Host "Fermeture du programme Excel [Terminé]"
LogTrace ("Fermeture du programme Excel [Terminé]")

Start-Sleep -s 5

Send_Mail; Write-Host "Envoi du mail avec les fichiers LOG et XLSX"
Write-Host -NoNewLine "FIN du script. Appuyer sur une touche pour quitter...`r`n"
$NULL = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")