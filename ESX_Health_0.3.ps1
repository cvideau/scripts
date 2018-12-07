### V0.3 (octobre 2018) - Développeur: Christophe VIDEAU
# Lien vers couleur Excel https://docs.microsoft.com/en-us/office/vba/images/colorin_za06050819.gif
# Lien vers couleur Write-Host https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-host?view=powershell-6
$WarningPreference = "SilentlyContinue"
Import-Module "VMware.VimAutomation.Core"
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
Add-PsSnapin VMware.VumAutomation
Add-PSSnapin -PassThru VMware.VimAutomation.Core
Clear-Host

Write-Host "Développement (2018) Christophe VIDEAU - Version 0.3`r`n" -ForegroundColor White

$Global:ElapsedTime_Start = 0; $Global:ElapsedTime_End = 0; $Global:ESX_Compliant = 0; $Global:ESX_NotCompliant = 0; $Global:ESX_NotCompliant_Item = 0; $Global:Cluster_NotCompliant_Item = 0
$Global:no_inc_ESX = 0; $Global:ESX_Counter = 0; $Global:Cluster_Counter = 0; $Global:Cluster_Inc = 0
$Global:Cluster = $NULL
$Global:ExcelLine_Start = 2		### Démarrage à la 2ème ligne du fichier Excel


### Fonction chargée de mesurer le temps de traitement
Function Get-ElapsedTime {
	$ElapsedTime_End = (Get-Date)
	$ElapsedTime = ($ElapsedTime_End - $ElapsedTime_Start).TotalSeconds
	$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
	$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
	If ($Sec.ToString().Length -eq 1) {	$Sec = "0" + $Sec	}
	Write-Host "[$($Min)min. $($Sec)sec]" -ForegroundColor White
}


### Fonction chargée de récupérer les valeurs du MATERIELS
Function Get-ESX_HARD { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "HARDWARE`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs HARDWARE")
	$ElapsedTime_Start = (Get-Date)
	
	$Esxcli2 = Get-ESXCLI -VMHost $ESX.Name -Server $vCenter
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 1) = $ESX.Name
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 2) = $Esxcli2.Hardware.Platform.Get.Invoke().VendorName
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 3) = $Esxcli2.Hardware.Platform.Get.Invoke().ProductName
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 4) = $Esxcli2.Hardware.Platform.Get.Invoke().SerialNumber
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 5) = $ESX.ExtensionData.Hardware.BiosInfo.BiosVersion
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 6) = $ESX.ExtensionData.Hardware.BiosInfo.ReleaseDate
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 7) = $Esxcli2.Storage.Core.Path.List.Invoke().AdapterTransportDetails | Where { $_.Device -eq "mpx.vmhba32:C0:T0:L0" }
	
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
	
	If ((Get-VmHostService -VMHost $ESX -Server $vCenter | Where-Object {$_.key -eq "ntpd"}).Running -eq "True") {	$ntpd_status = "Running"	} Else { $ntpd_status = "Not Running" }
	If ($Esxcli2.System.MaintenanceMode.Get.Invoke() -eq "Enabled") {	$ESX_State = "Maintenance"	} Else {	$ESX_State = "Connected"	}
	If ((Get-Compliance -Entity $ESX -Detailed -WarningAction "SilentlyContinue" | WHERE {$_.NotCompliantPatches -ne $NULL} | SELECT Status).Count -gt 0) {	$ComplianceLevel = "Baseline - Not compliant"	} Else {	$ComplianceLevel = "Baseline - Compliant"	}
	$Global:UTC = Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}
	$Global:UTC_Var1 = $UTC.Config.DateTimeInfo.TimeZone.Name
	$Global:UTC_Var2 = $UTC.Config.DateTimeInfo.TimeZone.GmtOffset
	$Global:Gateway = Get-VmHostNetwork -Host $ESX  -Server $vCenter | Select VMkernelGateway -ExpandProperty VirtualNic | Where {	$_.VMotionEnabled	} | Select -ExpandProperty VMkernelGateway -WarningAction "SilentlyContinue"
	
	New-VIProperty -Name EsxInstallDate -ObjectType VMHost -Value {	Param($ESX)
		$Esxcli = Get-Esxcli -VMHost $ESX.Name
		$Delta = [Convert]::ToInt64($esxcli.system.uuid.get.Invoke().Split('-')[0],16)
		(Get-Date -Year 1970 -Day 1 -Month 1 -Hour 0 -Minute 0 -Second 0).AddSeconds($delta)
	} -Force > $NULL
	$InstallDate = $(Get-VMHost -Name $vmhost | Select-Object -ExpandProperty EsxInstallDate)
	
	If ($oESX_NTP -eq $Gateway) {	$Global:Compliance_NTP = "Compliant"; $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 16).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 16).Font.Bold = $False	} Else {$Global:Compliance_NTP = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 16).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 16).Font.Bold = $True	}
	If ($UTC_Var2 -eq "0") {	$Global:Compliance_TimeOffset = "Compliant"	} Else {$Global:Compliance_TimeOffset = "Not Compliant"	}
	If ($ComplianceLevel -eq "Baseline - Compliant") {	$Global:Compliance_Level = "Compliant"	} Else {$Global:Compliance_Level = "Not Compliant"	}
	If ($NTPD_status -eq "Running") {	$Global:Compliance_NTPd = "Compliant"	} Else {$Global:Compliance_NTPd = "Not Compliant"	}
	
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 8) = "$DC"										# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 9) = "$Cluster"									# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 10) = $ESX_State								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 11) = $vCenter									# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 12) = "$oESXTAG"								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 13) = "PoweredOn"								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 14) = $ESX.Version + " - Build " + $ESX.Build
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 15) = "'" + $InstallDate
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 16) = "$oESX_NTP"
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 17) = $UTC_Var1 + "+" + $UTC_Var2				# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 18) = $Esxcli2.Hardware.Clock.Get.Invoke()		# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 19) = $ComplianceLevel							# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 20) = $NTPD_status

	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de la configuration SAN
Function Get-ESX_SAN { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "SAN`t`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs SAN")
	$ElapsedTime_Start = (Get-Date)
	$PathsRedondance_Error = 0
	
	Get-VMHostStorage -RescanAllHba -VMHost $ESX -Server $vCenter | Out-Null
	$Esxcli2 = Get-Esxcli -VMHost $ESX -Server $vCenter
	
	$Global:PathsSum = ($Esxcli2.storage.core.path.list.invoke() | Where {$_.Plugin -eq 'PowerPath'} | Select Device -Unique).Count
	$Global:PathsDead = ($Esxcli2.storage.core.path.list.invoke() | Where {($_.State -eq "dead") -and ($_.Plugin -eq 'PowerPath')} | Select Device -Unique).Count
	$Global:PathsActive = ($Esxcli2.storage.core.path.list.invoke() | Where {($_.State -eq "active") -and ($_.Plugin -eq 'PowerPath')} | Select Device -Unique).Count
	$Global:HBAAdapter = (Get-VMHostHba -VMHost $ESX -Type "FibreChannel").Count
	
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 21) = "$PathsSum" + " (" + $HBAAdapter + " adapters)"								# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 22) = $PathsDead																	# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 23) = $PathsActive																	# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 24) = $PathsRedondance + " (" + $PathsRedondance_Error + " chemin(s) en erreur)"	# Ligne, Colonne
	
	If ($PathsRedondance -eq "OK") {	$Global:Compliance_PathsDead = "Compliant"	} Else {	$Global:Compliance_PathsDead = "Not Compliant"	}

	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de  configuration ESXi
Function Get-ESX_HOST { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "HOTE`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs HOTE")
	$ElapsedTime_Start = (Get-Date)
	
	$Global:CurrentEVCMode = 	(Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Summary.CurrentEVCModeKey
	$Global:MaxEVCMode = 		(Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Summary.MaxEVCModeKey
	$Global:CPUPerformance = 	(Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Hardware.CpuPowerManagementInfo.CurrentPolicy
	$Global:HyperThreading = 	(Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Config.HyperThread.Active
	If ((Get-VmHostService -VMHost $ESX -Server $vCenter | Where-Object {$_.key -eq "vmware-fdm"}).Running -eq "True") {	$HAEnabled = "Running"	} Else { $HAEnabled = "Not Running" }
	$Global:Alarm = (Get-VMHost -Name $ESX).ExtensionData.AlarmActionsEnabled
	If ($Alarm -eq "True") { $Alarm = "Enabled" } Else { $Alarm = "Disabled" }
	If ((Get-AdvancedSetting -Entity $ESX -Name "Mem.AllocGuestLargePage").Value -eq "0") 	{ $LPages = "Disabled"	} 	Else { $LPages = "Enabled" }
	If ((Get-AdvancedSetting -Entity $ESX -Name "Mem.ShareForceSalting").Value -eq "0") 	{ $TPS = "Enabled" } 		Else { $TPS = "Disabled" }

	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 25) = $CPUPerformance											# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 26) = $HAEnabled												# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 27) = $HyperThreading											# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 28) = "Current: " + $CurrentEVCMode + ", Max: " + $MaxEVCMode	# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 29) = $Alarm													# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 30) = $TPS														# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 31) = $LPages													# Ligne, Colonne
	
	If ($CPUPerformance -eq "High Performance") {	$Global:Compliance_CPUPolicy = "Compliant"	} 	Else {	$Global:Compliance_CPUPolicy = "Not Compliant"	}
	If ((($HAEnabled -eq "Running") -and ($ESX_State -eq "Connected")) -or (($HAEnabled -eq "Not Running") -and ($ESX_State -eq "Maintenance"))) {	$Global:Compliance_HA = "Compliant"	} Else { $Global:Compliance_HA = "Not Compliant" }
	If ($HyperThreading -eq "VRAI")				{	$Global:Compliance_HT = "Compliant"	} 			Else {	$Global:Compliance_HT = "Not Compliant"	}
	If ($LPages -eq "Enabled") 					{	$Global:Compliance_LPages = "Compliant"	} 		Else {	$Global:Compliance_LPages = "Not Compliant"	}

	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de configuration du RESEAU
Function Get-ESX_NETWORK { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "RESEAUX`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs RESEAUX")
	$ElapsedTime_Start = (Get-Date)
	
	$Esxcli2 = Get-ESXCLI -VMHost $ESX -Server $vCenter
	$ESXvLAN = Get-VirtualPortGroup -VMHost $ESX -Server $vCenter -Distributed | Select Name
	ForEach ($DPortgroup in $ESXvLAN)	{
		$vLANID = ($ESXvLAN.Name).SubString(0, 4)
		Get-Unique -AsString -InputObject $vLANID | Out-Null
	}
	$vLANID = $vLANID -Replace "0000","" -Replace "_", "" -Replace "[^0-9]","" -Replace "  ",""
	$vLANID = "(" + $vLANID.Count + ")`n " + $vLANID
	
	$Global:ESXAdapter = $Esxcli2.Network.nic.list.Invoke()			| Where {$_.Link -eq "Up"}
	$Global:ESXAdapterName = $Esxcli2.Network.nic.list.Invoke()		| Where {$_.Link -eq "Up"} | Select Name, Speed, Duplex | Out-String
	$Global:ESXAdapterName = $ESXAdapterName -Replace "-","" -Replace "Name","" -Replace "Speed","" -Replace "Duplex","" -Replace "`r`n","" -Replace " ","" -Replace "10000", " 10000 " -Replace "vm", "`r`nvm"
	$Global:ESXAdapterCount = ($Esxcli2.Network.nic.list.Invoke() 	| Where {$_.Link -eq "Up"}).Count
	
	$vMotionIP = Get-VMHostNetworkAdapter -VMHost $ESX | Where {$_.DeviceName -eq "vmk1"} | Select IP | Out-String
	$vMotionIP = $vMotionIP -Replace "-","" -Replace "IP","" -Replace "`n","" -Replace " ",""
	$vMotionEnabled = (Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Summary.Config.VmotionEnabled
	If ($vMotionEnabled -eq $True) {	$vMotionEnabled = "vMotion enabled"	} Else {	$vMotionEnabled = "vMotion disabled"	}

	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 32) = $vLANID											# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 33) = "(" + $ESXAdapterCount + ") " + $ESXAdapterName	# Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 34) = $vMotionIP + " (" + $vMotionEnabled + ")"			# Ligne, Colonne
	
	If (($ESXAdapterName -notcontains "Half") -or ($ESXAdapterName -notcontains " 1000 ")) {	$Global:Compliance_NetworkFlow = "Compliant"	} Else {$Global:Compliance_Level = "Not Compliant"	}
	If ($vMotionEnabled -eq "vMotion enabled") {	$Global:Compliance_vMotion = "Compliant"	} Else {	$Global:Compliance_vMotion = "Not Compliant"	}

	Get-ElapsedTime
}


### Fonction chargée de récupérer les valeurs de configuration iLO
Function Get-HP_ILO { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs " -NoNewLine
	Write-Host "ILO`t`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Récupération des valeurs ILO")
	$ElapsedTime_Start = (Get-Date)
	
	$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 36) = $iLOversion # Ligne, Colonne
	
	Get-ElapsedTime
}


### Fonction chargée de vérifier si ESXi est conforme aux valeurs de base
Function Get-ESX_Compliant 	{
	Write-Host " * Analyse de la conformité " -NoNewLine
	Write-Host "ESXi`t`t" -NoNewLine -ForegroundColor White
	LogTrace (" * Analyse de la conformité ESXi")
	$ElapsedTime_Start = (Get-Date)
	
	If (($Global:Compliance_NTP 				-eq "Compliant") `
		-and ($Global:Compliance_TimeOffset 	-eq "Compliant") `
		-and ($Global:Compliance_Level 			-eq "Compliant") `
		-and ($Global:Compliance_NTPd 			-eq "Compliant") `
		-and ($Global:Compliance_PathsDead 		-eq "Compliant") `
		-and ($Global:Compliance_CPUPolicy 		-eq "Compliant") `
		-and ($Global:Compliance_HA 			-eq "Compliant") `
		-and ($Global:Compliance_HT 			-eq "Compliant") `
		-and ($Global:Compliance_LPages 		-eq "Compliant") `
		-and ($Global:Compliance_NetworkFlow 	-eq "Compliant") `
		-and ($Global:Compliance_vMotion 		-eq "Compliant")) { $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 37) = "[Conforme]"; $ESX_Compliant += 1	 }
	Else	{ $ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 37) = "[Non conforme]"; $ESX_NotCompliant += 1	}
	
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
	$ExcelLine = ($ExcelLine_Start..($ExcelLine_Start + $ESX_Counter - 1))
	ForEach($l in $ExcelLine)	{
		$ArrayTag		+= $ExcelWorkSheet.Cells.Item($l, 12).Text
		$ArrayVersion	+= $ExcelWorkSheet.Cells.Item($l, 14).Text
		$ArrayTimeZone 	+= $ExcelWorkSheet.Cells.Item($l, 17).Text
		$ArrayBaseline 	+= $ExcelWorkSheet.Cells.Item($l, 19).Text
		$ArrayNTPd	 	+= $ExcelWorkSheet.Cells.Item($l, 20).Text
		$ArraySAN 		+= $ExcelWorkSheet.Cells.Item($l, 21).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, 21).Text.IndexOf(" "))
		$ArraySANDeath	+= $ExcelWorkSheet.Cells.Item($l, 22).Text
		$ArraySANAlive	+= $ExcelWorkSheet.Cells.Item($l, 23).Text
		$ArraySANRedon	+= $ExcelWorkSheet.Cells.Item($l, 24).Text
		$ArrayCPUPolicy	+= $ExcelWorkSheet.Cells.Item($l, 25).Text
		$ArrayHA		+= $ExcelWorkSheet.Cells.Item($l, 26).Text
		$ArrayHyperTh	+= $ExcelWorkSheet.Cells.Item($l, 27).Text
		$ArrayEVC		+= $ExcelWorkSheet.Cells.Item($l, 28).Text
		$ArrayAlarme	+= $ExcelWorkSheet.Cells.Item($l, 29).Text
		$ArrayTPS 		+= $ExcelWorkSheet.Cells.Item($l, 30).Text
		$ArrayLPages 	+= $ExcelWorkSheet.Cells.Item($l, 31).Text
		$ArrayVLAN 		+= $ExcelWorkSheet.Cells.Item($l, 32).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, 32).Text.IndexOf(")") - 1)
		$ArrayLAN 		+= $ExcelWorkSheet.Cells.Item($l, 33).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, 33).Text.IndexOf(")") - 1)
	}
	Get-ElapsedTime

	### Détermination des valeurs de références (majoritaires)
	Write-Host " * Détermination des valeurs de références`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Détermination des valeurs de références")
	$ElapsedTime_Start = (Get-Date)

	$Ref_ArrayTag 		= ($ArrayTag 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayVersion	= ($ArrayVersion 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayTimeZone 	= ($ArrayTimeZone 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayBaseline 	= ($ArrayBaseline 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayNTPd	 	= ($ArrayNTPd 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArraySAN 		= ($ArraySAN 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArraySANDeath	= ($ArraySANDeath	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArraySANAlive	= ($ArraySANAlive	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArraySANRedon	= ($ArraySANRedon	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayCPUPolicy	= ($ArrayCPUPolicy 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayHA		= ($ArrayHA 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayHyperTh	= ($ArrayHyperTh 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayEVC		= ($ArrayEVC 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayAlarme	= ($ArrayAlarme 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayTPS 		= ($ArrayTPS		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayLPages	= ($ArrayLPages 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayVLAN 		= ($ArrayVLAN 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayLAN 		= ($ArrayLAN 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	
	Get-ElapsedTime
	
	### Inscription des valeurs de références par cluster dans la 2ème feuille du fichier Excel
	Write-Host " * Inscription des valeurs de références dans Excel`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Inscription des valeurs de références dans Excel")
	$ElapsedTime_Start = (Get-Date)
	
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 1)		= $Cluster.Name
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 2)		= $Ref_ArrayVersion
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 4) 	= $Ref_ArrayTimeZone
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 5) 	= $Ref_ArrayBaseline
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 6) 	= $Ref_ArrayNTPd
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 7) 	= $Ref_ArraySAN
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 8) 	= $Ref_ArrayCPUPolicy
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 9) 	= $Ref_ArrayHA
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 10) 	= $Ref_ArrayHyperTh
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 11) 	= $Ref_ArrayEVC
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 12) 	= $Ref_ArrayAlarme
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 13) 	= $Ref_ArrayTPS
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 14) 	= $Ref_ArrayLPages
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 15) 	= $Ref_ArrayVLAN
	$ExcelWorkSheet_Ref.Cells.Item($Cluster_Inc + 1, 16) 	= $Ref_ArrayLAN
	
	Get-ElapsedTime

	### Vérification des valeurs par rapport aux valeurs majoritaires
	Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Vérification des cellules vs valeurs majoritaires")
	$ElapsedTime_Start = (Get-Date)
	
	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = ($ExcelLine_Start..($ExcelLine_Start + $ESX_Counter - 1))
	ForEach($l in $ExcelLine)	{
		If ($ExcelWorkSheet.Cells.Item($l, 12).Text -ne $Ref_ArrayTag) 			{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ Tag"; 					$ExcelWorkSheet.Cells.Item($l, 12).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 12).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 12).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 12).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 14).Text -ne $Ref_ArrayVersion) 		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ version"; 				$ExcelWorkSheet.Cells.Item($l, 14).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 14).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 14).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 14).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 17).Text -ne $Ref_ArrayTimeZone) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ TimeZone"; 			$ExcelWorkSheet.Cells.Item($l, 17).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 17).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 17).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 17).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 19).Text -ne $Ref_ArrayBaseline) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ Baseline"; 			$ExcelWorkSheet.Cells.Item($l, 19).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 19).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 19).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 19).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 20).Text -ne $Ref_ArrayNTPd) 		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ NTP démon"; 			$ExcelWorkSheet.Cells.Item($l, 20).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 20).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 20).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 20).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 21).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, 21).Text.IndexOf(" ")) -ne $Ref_ArraySAN) {	$ExcelWorkSheet.Cells.Item($l, 37)	= 				$ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ chemins SAN"; $ExcelWorkSheet.Cells.Item($l, 21).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 21).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 21).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 21).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 22).Text -ne $Ref_ArraySANDeath) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", chemin(s) mort(s)"; 		$ExcelWorkSheet.Cells.Item($l, 22).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 22).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 22).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 22).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 23).Text -ne $Ref_ArraySANAlive) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", chemin(s) présent(s)"; 	$ExcelWorkSheet.Cells.Item($l, 23).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 23).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 23).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 23).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 24).Text -ne $Ref_ArraySANRedon) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur redondance SAN"; 	$ExcelWorkSheet.Cells.Item($l, 24).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 24).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 24).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 24).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 25).Text -ne $Ref_ArrayCPUPolicy) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ CPU Policy"; 			$ExcelWorkSheet.Cells.Item($l, 25).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 25).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 25).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 25).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 26).Text -ne $Ref_ArrayHA) 			{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ HA"; 					$ExcelWorkSheet.Cells.Item($l, 26).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 26).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 26).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 26).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 27).Text -ne $Ref_ArrayHyperTh) 		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ HyperThreading"; 		$ExcelWorkSheet.Cells.Item($l, 27).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 27).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 27).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 27).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 28).Text -ne $Ref_ArrayEVC) 			{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ EVC"; 					$ExcelWorkSheet.Cells.Item($l, 28).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 28).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 28).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 28).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 29).Text -ne $Ref_ArrayAlarme)		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ alarmes"; 				$ExcelWorkSheet.Cells.Item($l, 29).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 29).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 29).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 29).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 30).Text -ne $Ref_ArrayTPS) 			{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ TPS"; 					$ExcelWorkSheet.Cells.Item($l, 30).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 30).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 30).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 30).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 31).Text -ne $Ref_ArrayLPages) 		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ Large Pages"; 			$ExcelWorkSheet.Cells.Item($l, 31).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 31).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 31).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 31).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 32).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, 32).Text.IndexOf(")") - 1) -ne $Ref_ArrayVLAN) {	$ExcelWorkSheet.Cells.Item($l, 37)	= 		$ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ VLAN"; $ExcelWorkSheet.Cells.Item($l, 32).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 32).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 32).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 32).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 33).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, 33).Text.IndexOf(")") - 1) -ne $Ref_ArrayLAN) {	$ExcelWorkSheet.Cells.Item($l, 37)	= 			$ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ adaptateurs LAN"; $ExcelWorkSheet.Cells.Item($l, 33).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 33).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 33).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 33).Font.Bold = $False	}
	}
	
	Get-ElapsedTime
	Write-Host "Vérification de l'homogénéité des données du cluster..." -NoNewLine
	Write-Host "(Terminée)" -ForegroundColor Green
	Write-Host "En résumé: $ESX_NotCompliant ESXi non conforme(s), $ESX_NotCompliant_Item différence(s)`r`n" -ForegroundColor White
	LogTrace ("Vérification de l'homogénéité des données du cluster... (Terminée)")
}


### Fonction chargée de la comparaison/l'homogénéité des valeurs entre elles et relatives aux ESXi ciblés
Function Get-ESX_Compare_Part {
	Write-Host "Vérification de l'homogénéité des ESXi ciblés... " -NoNewLine
	Write-Host "En cours" -ForegroundColor Green
	Write-Host "RAPPEL: $Mode [ID $Mode_Var]" -ForegroundColor Red
	LogTrace ("Vérification de l'homogénéité des ESXi ciblés...")
	
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
	$ExcelLine = ($ExcelLine_Start..($ExcelLine_Start + $ESX_Counter - 1))
	ForEach($l in $ExcelLine)	{
		$ArrayTag		+= $ExcelWorkSheet.Cells.Item($l, 12).Text
		$ArrayVersion	+= $ExcelWorkSheet.Cells.Item($l, 14).Text
		$ArrayTimeZone 	+= $ExcelWorkSheet.Cells.Item($l, 17).Text
		$ArrayBaseline 	+= $ExcelWorkSheet.Cells.Item($l, 19).Text
		$ArrayNTPd	 	+= $ExcelWorkSheet.Cells.Item($l, 20).Text
		$ArraySAN 		+= $ExcelWorkSheet.Cells.Item($l, 21).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, 21).Text.IndexOf(" "))
		$ArraySANDeath	+= $ExcelWorkSheet.Cells.Item($l, 22).Text
		$ArraySANAlive	+= $ExcelWorkSheet.Cells.Item($l, 23).Text
		$ArraySANRedon	+= $ExcelWorkSheet.Cells.Item($l, 24).Text
		$ArrayCPUPolicy	+= $ExcelWorkSheet.Cells.Item($l, 25).Text
		$ArrayHA		+= $ExcelWorkSheet.Cells.Item($l, 26).Text
		$ArrayHyperTh	+= $ExcelWorkSheet.Cells.Item($l, 27).Text
		$ArrayEVC		+= $ExcelWorkSheet.Cells.Item($l, 28).Text
		$ArrayAlarme	+= $ExcelWorkSheet.Cells.Item($l, 29).Text
		$ArrayTPS 		+= $ExcelWorkSheet.Cells.Item($l, 30).Text
		$ArrayLPages 	+= $ExcelWorkSheet.Cells.Item($l, 31).Text
		$ArrayVLAN 		+= $ExcelWorkSheet.Cells.Item($l, 32).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, 32).Text.IndexOf(")") - 1)
		$ArrayLAN 		+= $ExcelWorkSheet.Cells.Item($l, 33).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, 33).Text.IndexOf(")") - 1)
	}
	Get-ElapsedTime

	### Détermination des valeurs de références (majoritaires)
	Write-Host " * Détermination des valeurs de références`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Détermination des valeurs de références")
	$ElapsedTime_Start = (Get-Date)

	$Ref_ArrayTag 		= ($ArrayTag 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayVersion	= ($ArrayVersion 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayTimeZone 	= ($ArrayTimeZone 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayBaseline 	= ($ArrayBaseline 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayNTPd	 	= ($ArrayNTPd 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArraySAN 		= ($ArraySAN 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArraySANDeath	= ($ArraySANDeath	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArraySANAlive	= ($ArraySANAlive	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArraySANRedon	= ($ArraySANRedon	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayCPUPolicy	= ($ArrayCPUPolicy 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayHA		= ($ArrayHA 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayHyperTh	= ($ArrayHyperTh 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayEVC		= ($ArrayEVC 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayAlarme	= ($ArrayAlarme 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayTPS 		= ($ArrayTPS		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayLPages	= ($ArrayLPages 	| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayVLAN 		= ($ArrayVLAN 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayLAN 		= ($ArrayLAN 		| Group | ?{	$_.Count -gt ($ESX_Counter/2)	}).Values | Select -Last 1
	
	Get-ElapsedTime
	
	### Vérification des valeurs par rapport aux valeurs majoritaires
	Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Vérification des cellules vs valeurs majoritaires")
	$ElapsedTime_Start = (Get-Date)
	
	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = ($ExcelLine_Start..($ExcelLine_Start + $ESX_Counter - 1))
	ForEach($l in $ExcelLine)	{
		If ($ExcelWorkSheet.Cells.Item($l, 12).Text -ne $Ref_ArrayTag) 			{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ Tag"; 					$ExcelWorkSheet.Cells.Item($l, 12).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 12).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 12).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 12).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 14).Text -ne $Ref_ArrayVersion) 		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ version"; 				$ExcelWorkSheet.Cells.Item($l, 14).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 14).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 14).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 14).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 17).Text -ne $Ref_ArrayTimeZone) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ TimeZone"; 			$ExcelWorkSheet.Cells.Item($l, 17).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 17).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 17).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 17).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 19).Text -ne $Ref_ArrayBaseline) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ Baseline"; 			$ExcelWorkSheet.Cells.Item($l, 19).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 19).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 19).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 19).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 20).Text -ne $Ref_ArrayNTPd) 		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ NTP démon"; 			$ExcelWorkSheet.Cells.Item($l, 20).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 20).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 20).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 20).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 21).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, 21).Text.IndexOf(" ")) -ne $Ref_ArraySAN) {	$ExcelWorkSheet.Cells.Item($l, 37)	=				$ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ chemins SAN"; $ExcelWorkSheet.Cells.Item($l, 21).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 21).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 21).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 21).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 22).Text -ne $Ref_ArraySANDeath) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", chemin(s) mort(s)"; 		$ExcelWorkSheet.Cells.Item($l, 22).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 22).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 22).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 22).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 23).Text -ne $Ref_ArraySANAlive) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", chemin(s) présent(s)"; 	$ExcelWorkSheet.Cells.Item($l, 23).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 23).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 23).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 23).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 24).Text -ne $Ref_ArraySANRedon) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur redondance SAN"; 	$ExcelWorkSheet.Cells.Item($l, 24).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 24).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 24).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 24).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 25).Text -ne $Ref_ArrayCPUPolicy) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ CPU Policy"; 			$ExcelWorkSheet.Cells.Item($l, 25).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 25).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 25).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 25).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 26).Text -ne $Ref_ArrayHA) 			{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ HA"; 					$ExcelWorkSheet.Cells.Item($l, 26).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 26).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 26).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 26).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 27).Text -ne $Ref_ArrayHyperTh) 		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ HyperThreading"; 		$ExcelWorkSheet.Cells.Item($l, 27).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 27).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 27).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 27).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 28).Text -ne $Ref_ArrayEVC) 			{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ EVC"; 					$ExcelWorkSheet.Cells.Item($l, 28).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 28).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 28).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 28).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 29).Text -ne $Ref_ArrayAlarme)		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ alarmes"; 				$ExcelWorkSheet.Cells.Item($l, 29).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 29).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 29).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 29).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 30).Text -ne $Ref_ArrayTPS) 			{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ TPS"; 					$ExcelWorkSheet.Cells.Item($l, 30).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 30).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 30).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 30).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 31).Text -ne $Ref_ArrayLPages) 		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ Large Pages"; 			$ExcelWorkSheet.Cells.Item($l, 31).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 31).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 31).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 31).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 32).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, 32).Text.IndexOf(")") - 1) -ne $Ref_ArrayVLAN) {	$ExcelWorkSheet.Cells.Item($l, 37)	= 		$ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ VLAN"; $ExcelWorkSheet.Cells.Item($l, 32).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 32).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 32).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 32).Font.Bold = $False	}
		If ($ExcelWorkSheet.Cells.Item($l, 33).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, 33).Text.IndexOf(")") - 1) -ne $Ref_ArrayLAN) {	$ExcelWorkSheet.Cells.Item($l, 37)	= 			$ExcelWorkSheet.Cells.Item($l, 37).Text + ", ≠ adaptateurs LAN"; $ExcelWorkSheet.Cells.Item($l, 33).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 33).Font.Bold = $True; $ESX_NotCompliant_Item += 1	} Else {	$ExcelWorkSheet.Cells.Item($l, 33).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 33).Font.Bold = $False	}
	}
	
	Get-ElapsedTime
	Write-Host "Vérification de l'homogénéité des données ESXi ciblés..." -NoNewLine
	Write-Host "(Terminée)" -ForegroundColor Green
	Write-Host "En résumé: $ESX_NotCompliant ESXi non conforme(s), $ESX_NotCompliant_Item différence(s)`r`n" -ForegroundColor White
	LogTrace ("Vérification de l'homogénéité des données du cluster... (Terminée)")
}


### Fonction chargée de la comparaison/l'homogénéité des valeurs avec celles du fichier de référence
Function Get-ESX_Compare_Part_Ref {
	Write-Host "Comparaison des ESX ciblés avec les références clusters... " -NoNewLine
	Write-Host "En cours" -ForegroundColor Green
	LogTrace ("Comparaison des ESX ciblés avec les références clusters...")
	
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

	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = ($ExcelLine_Start..($ExcelLine_Start + $ESX_Counter - 1))

	ForEach($l in $ExcelLine)	{
		### Recherche du cluster relatif à l'ESXi ciblé dans la feuille de référence Excel
		$GetName = $ExcelWorkSheet_Ref.Range("A1").EntireColumn.Find($ExcelWorkSheet.Cells.Item($l, 9))
		If (!($($GetName.Row))) {

			$Ref_ArrayVersion	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 2).Text		# Ligne, Colonne
			$Ref_ArrayTimeZone 	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 4).Text		# Ligne, Colonne
			$Ref_ArrayBaseline 	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 5).Text		# Ligne, Colonne
			$Ref_ArrayNTPd	 	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 6).Text		# Ligne, Colonne
			$Ref_ArraySAN		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 7).Text		# Ligne, Colonne
			$Ref_ArrayCPUPolicy	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 8).Text		# Ligne, Colonne
			$Ref_ArrayHA		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 9).Text		# Ligne, Colonne
			$Ref_ArrayHyperTh	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 10).Text		# Ligne, Colonne
			$Ref_ArrayEVC		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 11).Text		# Ligne, Colonne
			$Ref_ArrayAlarme	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 12).Text		# Ligne, Colonne
			$Ref_ArrayTPS 		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 13).Text		# Ligne, Colonne
			$Ref_ArrayLPages 	= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 14).Text		# Ligne, Colonne
			$Ref_ArrayVLAN		= $ExcelWorkSheet_Ref.Cells.Item($GetName.Row, 15).Text		# Ligne, Colonne

			Get-ElapsedTime
		
			### Vérification des données par rapport aux valeurs majoritaires
			Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
			LogTrace (" * Vérification des cellules vs valeurs majoritaires")
			$ElapsedTime_Start = (Get-Date)
		
			If ($ExcelWorkSheet.Cells.Item($l, 14).Text -ne $Ref_ArrayVersion) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur version"; 		$ExcelWorkSheet.Cells.Item($l,14).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 14).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	}	Else {	$ExcelWorkSheet.Cells.Item($l, 14).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 14).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 17).Text -ne $Ref_ArrayTimeZone)	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur TimeZone"; 		$ExcelWorkSheet.Cells.Item($l, 17).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 17).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	}	Else {	$ExcelWorkSheet.Cells.Item($l, 17).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 17).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 19).Text -ne $Ref_ArrayBaseline) {	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur Baseline"; 		$ExcelWorkSheet.Cells.Item($l, 19).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 19).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	}	Else {	$ExcelWorkSheet.Cells.Item($l, 19).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 19).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 20).Text -ne $Ref_ArrayNTPd) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur NTP démon"; 		$ExcelWorkSheet.Cells.Item($l, 20).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 20).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	}	Else {	$ExcelWorkSheet.Cells.Item($l, 20).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 20).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 21).Text.Substring(0, $ExcelWorkSheet.Cells.Item($l, 21).Text.IndexOf(" ")) -ne $Ref_ArraySAN)		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur chemins SAN"; 	$ExcelWorkSheet.Cells.Item($l, 21).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 21).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	}	Else {	$ExcelWorkSheet.Cells.Item($l, 21).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 21).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 25).Text -ne $Ref_ArrayCPUPolicy){	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur CPU Policy"; 		$ExcelWorkSheet.Cells.Item($l,25).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 25).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 	Else {	$ExcelWorkSheet.Cells.Item($l, 25).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 25).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 26).Text -ne $Ref_ArrayHA) 		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur HA"; 				$ExcelWorkSheet.Cells.Item($l, 26).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 26).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 	Else {	$ExcelWorkSheet.Cells.Item($l, 26).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 26).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 27).Text -ne $Ref_ArrayHyperTh) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur HyperThreading"; 	$ExcelWorkSheet.Cells.Item($l, 27).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 27).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	}	Else {	$ExcelWorkSheet.Cells.Item($l, 27).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 27).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 28).Text -ne $Ref_ArrayEVC) 		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur EVC"; 			$ExcelWorkSheet.Cells.Item($l, 28).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 28).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 	Else {	$ExcelWorkSheet.Cells.Item($l, 28).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 28).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 29).Text -ne $Ref_ArrayAlarme)	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur alarmes"; 		$ExcelWorkSheet.Cells.Item($l, 29).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 29).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	}	Else {	$ExcelWorkSheet.Cells.Item($l, 29).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 29).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 30).Text -ne $Ref_ArrayTPS) 		{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur TPS"; 			$ExcelWorkSheet.Cells.Item($l, 30).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 30).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	}	Else {	$ExcelWorkSheet.Cells.Item($l, 30).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 30).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 31).Text -ne $Ref_ArrayLPages) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur Large Pages"; 	$ExcelWorkSheet.Cells.Item($l, 31).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 31).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	}	Else {	$ExcelWorkSheet.Cells.Item($l, 31).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 31).Font.Bold = $False	}
			If ($ExcelWorkSheet.Cells.Item($l, 32).Text.Substring(1, $ExcelWorkSheet.Cells.Item($l, 32).Text.IndexOf(")") - 1) -ne $Ref_ArrayVLAN) 	{	$ExcelWorkSheet.Cells.Item($l, 37)	= $ExcelWorkSheet.Cells.Item($l, 37).Text + ", erreur VLAN"; 			$ExcelWorkSheet.Cells.Item($l, 32).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($l, 32).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	}	Else {	$ExcelWorkSheet.Cells.Item($l, 32).Font.ColorIndex = 1; $ExcelWorkSheet.Cells.Item($l, 32).Font.Bold = $False	}
		}
	}
	
	Get-ElapsedTime
	
	If (!($($GetName.Row))) {
		Write-Host "Comparaison des ESX ciblés avec les références clusters... (Terminée)"
		Write-Host "En résumé:`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * Cluster trouvé à la ligne: $($GetName.Row)`r`n * ESXi non conforme(s): $ESX_NotCompliant`r`n * Différence(s): $ESX_NotCompliant_Item`r`n" -ForegroundColor White
		LogTrace ("Comparaison des ESX ciblés avec les références clusters...... (Terminée)")
	}
	Else {
		Write-Host "Comparaison des ESX ciblés impossible, cluster non trouvé... (Terminée)"
		Write-Host "En résumé:`r`n * vCenter: $vCenter`r`n * Cluster: $Cluster`r`n * ESXi non conforme(s): $ESX_NotCompliant`r`n * Différence(s): $ESX_NotCompliant_Item`r`n" -ForegroundColor White
		LogTrace ("Comparaison des ESX ciblés impossible, cluster non trouvé... (Terminée)")
	}
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

	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = ($ExcelLine_Start..($ExcelLine_Start + $ESX_Counter - 1))
	
	ForEach($l in $ExcelLine)	{
		$ArrayVersion	+= $ExcelWorkSheet_Ref.Cells.Item($l, 2).Text
		$ArrayTimeZone 	+= $ExcelWorkSheet_Ref.Cells.Item($l, 4).Text
		$ArrayBaseline 	+= $ExcelWorkSheet_Ref.Cells.Item($l, 5).Text
		$ArrayNTPd	 	+= $ExcelWorkSheet_Ref.Cells.Item($l, 6).Text
		$ArrayCPUPolicy	+= $ExcelWorkSheet_Ref.Cells.Item($l, 8).Text
		$ArrayHA		+= $ExcelWorkSheet_Ref.Cells.Item($l, 9).Text
		$ArrayHyperTh	+= $ExcelWorkSheet_Ref.Cells.Item($l, 10).Text
		$ArrayEVC		+= $ExcelWorkSheet_Ref.Cells.Item($l, 11).Text
		$ArrayAlarme	+= $ExcelWorkSheet_Ref.Cells.Item($l, 12).Text
		$ArrayTPS 		+= $ExcelWorkSheet_Ref.Cells.Item($l, 13).Text
		$ArrayLPages 	+= $ExcelWorkSheet_Ref.Cells.Item($l, 14).Text
	}
	Get-ElapsedTime

	### Détermination des valeurs de références (majoritaires)
	Write-Host " * Détermination des valeurs de références`t`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Détermination des valeurs de références")
	$ElapsedTime_Start = (Get-Date)

	$Ref_ArrayVersion	= ($ArrayVersion 	| Group | ?{	$_.Count -gt ($Cluster_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayTimeZone 	= ($ArrayTimeZone 	| Group | ?{	$_.Count -gt ($Cluster_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayBaseline 	= ($ArrayBaseline 	| Group | ?{	$_.Count -gt ($Cluster_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayNTPd	 	= ($ArrayNTPd 		| Group | ?{	$_.Count -gt ($Cluster_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayCPUPolicy	= ($ArrayCPUPolicy 	| Group | ?{	$_.Count -gt ($Cluster_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayHA		= ($ArrayHA 		| Group | ?{	$_.Count -gt ($Cluster_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayHyperTh	= ($ArrayHyperTh 	| Group | ?{	$_.Count -gt ($Cluster_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayEVC		= ($ArrayEVC 		| Group | ?{	$_.Count -gt ($Cluster_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayAlarme	= ($ArrayAlarme 	| Group | ?{	$_.Count -gt ($Cluster_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayTPS 		= ($ArrayTPS		| Group | ?{	$_.Count -gt ($Cluster_Counter/2)	}).Values | Select -Last 1
	$Ref_ArrayLPages	= ($ArrayLPages 	| Group | ?{	$_.Count -gt ($Cluster_Counter/2)	}).Values | Select -Last 1
	
	Get-ElapsedTime
	
	### Vérification des données par rapport aux valeurs majoritaires
	Write-Host " * Vérification des cellules vs valeurs majoritaires`t" -NoNewLine -ForegroundColor DarkYellow
	LogTrace (" * Vérification des cellules vs valeurs majoritaires")
	$ElapsedTime_Start = (Get-Date)
	
	# Boucle déterminant la première ligne Excel du cluster à la dernière
	$ExcelLine = ($ExcelLine_Start..($ExcelLine_Start + $ESX_Counter - 1))
	ForEach($l in $ExcelLine)	{
		If ($ExcelWorkSheet_Ref.Cells.Item($l, 2).Text -ne $Ref_ArrayVersion) 	{	$ExcelWorkSheet_Ref.Cells.Item($l, 17)	= $ExcelWorkSheet_Ref.Cells.Item($l, 17).Text + ", erreur version"; $ExcelWorkSheet_Ref.Cells.Item($l, 2).Font.ColorIndex = 3; $ExcelWorkSheet_Ref.Cells.Item($l, 2).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 				Else {	$ExcelWorkSheet_Ref.Cells.Item($l, 2).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, 2).Font.Bold = $False	}
		If ($ExcelWorkSheet_Ref.Cells.Item($l, 4).Text -ne $Ref_ArrayTimeZone) 	{	$ExcelWorkSheet_Ref.Cells.Item($l, 17)	= $ExcelWorkSheet_Ref.Cells.Item($l, 17).Text + ", erreur TimeZone"; $ExcelWorkSheet_Ref.Cells.Item($l, 4).Font.ColorIndex = 3; $ExcelWorkSheet_Ref.Cells.Item($l, 4).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 			Else {	$ExcelWorkSheet_Ref.Cells.Item($l, 4).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, 4).Font.Bold = $False	}
		If ($ExcelWorkSheet_Ref.Cells.Item($l, 5).Text -ne $Ref_ArrayBaseline) 	{	$ExcelWorkSheet_Ref.Cells.Item($l, 17)	= $ExcelWorkSheet_Ref.Cells.Item($l, 17).Text + ", erreur Baseline"; $ExcelWorkSheet_Ref.Cells.Item($l, 5).Font.ColorIndex = 3; $ExcelWorkSheet_Ref.Cells.Item($l, 5).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 			Else {	$ExcelWorkSheet_Ref.Cells.Item($l, 5).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, 5).Font.Bold = $False	}
		If ($ExcelWorkSheet_Ref.Cells.Item($l, 6).Text -ne $Ref_ArrayNTPd) 		{	$ExcelWorkSheet_Ref.Cells.Item($l, 17)	= $ExcelWorkSheet_Ref.Cells.Item($l, 17).Text + ", erreur NTP démon"; $ExcelWorkSheet_Ref.Cells.Item($l, 6).Font.ColorIndex = 3; $ExcelWorkSheet_Ref.Cells.Item($l, 6).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 			Else {	$ExcelWorkSheet_Ref.Cells.Item($l, 6).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, 6).Font.Bold = $False	}
		If ($ExcelWorkSheet_Ref.Cells.Item($l, 8).Text -ne $Ref_ArrayCPUPolicy) {	$ExcelWorkSheet_Ref.Cells.Item($l, 17)	= $ExcelWorkSheet_Ref.Cells.Item($l, 17).Text + ", erreur CPU Policy"; $ExcelWorkSheet_Ref.Cells.Item($l, 8).Font.ColorIndex = 3; $ExcelWorkSheet_Ref.Cells.Item($l, 8).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 			Else {	$ExcelWorkSheet_Ref.Cells.Item($l, 8).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, 8).Font.Bold = $False	}
		If ($ExcelWorkSheet_Ref.Cells.Item($l, 9).Text -ne $Ref_ArrayHA) 		{	$ExcelWorkSheet_Ref.Cells.Item($l, 17)	= $ExcelWorkSheet_Ref.Cells.Item($l, 17).Text + ", erreur HA"; $ExcelWorkSheet_Ref.Cells.Item($l, 9).Font.ColorIndex = 3; $ExcelWorkSheet_Ref.Cells.Item($l, 9).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 					Else {	$ExcelWorkSheet_Ref.Cells.Item($l, 9).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, 9).Font.Bold = $False	}
		If ($ExcelWorkSheet_Ref.Cells.Item($l, 10).Text -ne $Ref_ArrayHyperTh) 	{	$ExcelWorkSheet_Ref.Cells.Item($l, 17)	= $ExcelWorkSheet_Ref.Cells.Item($l, 17).Text + ", erreur HyperThreading"; $ExcelWorkSheet_Ref.Cells.Item($l, 10).Font.ColorIndex = 3; $ExcelWorkSheet_Ref.Cells.Item($l, 10).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 	Else {	$ExcelWorkSheet_Ref.Cells.Item($l, 10).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, 10).Font.Bold = $False	}
		If ($ExcelWorkSheet_Ref.Cells.Item($l, 11).Text -ne $Ref_ArrayEVC) 		{	$ExcelWorkSheet_Ref.Cells.Item($l, 17)	= $ExcelWorkSheet_Ref.Cells.Item($l, 17).Text + ", erreur EVC"; $ExcelWorkSheet_Ref.Cells.Item($l, 11).Font.ColorIndex = 3; $ExcelWorkSheet_Ref.Cells.Item($l, 11).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 				Else {	$ExcelWorkSheet_Ref.Cells.Item($l, 11).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, 11).Font.Bold = $False	}
		If ($ExcelWorkSheet_Ref.Cells.Item($l, 12).Text -ne $Ref_ArrayAlarme)	{	$ExcelWorkSheet_Ref.Cells.Item($l, 17)	= $ExcelWorkSheet_Ref.Cells.Item($l, 17).Text + ", erreur alarmes"; $ExcelWorkSheet_Ref.Cells.Item($l, 12).Font.ColorIndex = 3; $ExcelWorkSheet_Ref.Cells.Item($l, 12).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 			Else {	$ExcelWorkSheet_Ref.Cells.Item($l, 12).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, 12).Font.Bold = $False	}
		If ($ExcelWorkSheet_Ref.Cells.Item($l, 13).Text -ne $Ref_ArrayTPS) 		{	$ExcelWorkSheet_Ref.Cells.Item($l, 17)	= $ExcelWorkSheet_Ref.Cells.Item($l, 17).Text + ", erreur TPS"; $ExcelWorkSheet_Ref.Cells.Item($l, 13).Font.ColorIndex = 3; $ExcelWorkSheet_Ref.Cells.Item($l, 13).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 				Else {	$ExcelWorkSheet_Ref.Cells.Item($l, 13).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, 13).Font.Bold = $False	}
		If ($ExcelWorkSheet_Ref.Cells.Item($l, 14).Text -ne $Ref_ArrayLPages) 	{	$ExcelWorkSheet_Ref.Cells.Item($l, 17)	= $ExcelWorkSheet_Ref.Cells.Item($l, 17).Text + ", erreur Large Pages"; $ExcelWorkSheet_Ref.Cells.Item($l, 14).Font.ColorIndex = 3; $ExcelWorkSheet_Ref.Cells.Item($l, 14).Font.Bold = $True; $Cluster_NotCompliant_Item += 1	} 		Else {	$ExcelWorkSheet_Ref.Cells.Item($l, 14).Font.ColorIndex = 1; $ExcelWorkSheet_Ref.Cells.Item($l, 14).Font.Bold = $False	}
	}
	
	Get-ElapsedTime
	Write-Host "Vérification de l'homogénéité des clusters par ENV... (Terminée)"
	Write-Host "En résumé: $ESX_NotCompliant ESXi non conforme(s), $ESX_NotCompliant_Item différence(s)`r`n" -ForegroundColor White
	LogTrace ("Vérification de l'homogénéité des clusters par ENV... (Terminée)")
}


### Fonction chargée de l'envoi de messages électroniques (En fin de traitement)
Function Send_Mail {
	$Sender1 = "Christophe.VIDEAU-ext@ca-ts.fr"
	$Sender2 = #"MCO.Infra.OS.distribues@ca-ts.fr"
	$From = "[Conformité ESXi] Weekly check <ESXiCompliance.report@ca-ts.fr>"
	$Subject = "[Conformité ESXi] Compte-rendu operationnel {Conformité infrastructures VMware}"
	If ($BodyMail_Error -ne $NULL) {	$Body = "Bonjour,<BR><BR>Vous trouverez en pi&egrave;ce jointe le fichier LOG et la synth&egrave;se technique de l&rsquo;op&eacute;ration.<BR><BR><U>En bref:</U><BR>Mode: " + $Mode + "<BR>Nombre d'ESXi conformes:<B><I> " + $ESX_Compliant + "</I></B><BR>Nombre d'ESXi non conforme(s):<B><I> " + $ESX_NotCompliant + " (" + $ESX_NotCompliant_Item + " différences depuis le début)</I><BR>Nombre de clusters non conforme(s):<B><I> " + $Cluster_NotCompliant + " (" + $Cluster_NotCompliant_Item + " différences)</I></B><BR>Dur&eacute;e du traitement: <I>" + $($Min) + "min. " + $($Sec) + "sec.</I><BR><BR>----------------------" + $BodyMail_Error + "<BR>----------------------<BR><BR><BR>Cordialement.<BR>L'&eacute;quipe d&eacute;veloppement (Contact: Hugues de TERSSAC)<BR><BR>Ce message a &eacute;t&eacute; envoy&eacute; par un automate, ne pas y r&eacute;pondre."	}
	Else {	$Body = "Bonjour,<BR><BR>Vous trouverez en pi&egrave;ce jointe le fichier LOG et la synth&egrave;se technique de l&rsquo;op&eacute;ration.<BR><BR><U>En bref:</U><BR>Mode: " + $Mode + "<BR>Nombre d'ESXi conformes:<B><I> " + $ESX_Compliant + "</I></B><BR>Nombre d'ESXi non conforme(s):<B><I> " + $ESX_NotCompliant + " (" + $ESX_NotCompliant_Item + " différences depuis le début)</I></B><BR>Nombre de clusters non conforme(s):<B><I> " + $Cluster_NotCompliant + " (" + $Cluster_NotCompliant_Item + " différences)</I><BR>Dur&eacute;e du traitement: <I>" + $($Min) + "min. " + $($Sec) + "sec.</I><BR><BR>Cordialement.<BR>L'&eacute;quipe d&eacute;veloppement (Contact: Hugues de TERSSAC)<BR><BR>Ce message a &eacute;t&eacute; envoy&eacute; par un automate, ne pas y r&eacute;pondre."	}
	$Attachments = $FicLog, $FicRes
	$SMTP = "muz10-e1smtp-IN-DC-INT.zres.ztech"
	Send-MailMessage -To $Sender1, $Sender2 -From $From -Subject $Subject -Body $Body -Attachments $Attachments -SmtpServer $SMTP -Priority High -BodyAsHTML
}


### Fonction chargée du vidage de feuille Excel
Function Clear_Excel_Feuille { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$Feuille )
	For($i = 2 ; $i -eq 10000 ; $i++)	{
		If ($Feuille.Cells.Item($i, 1).Text) {	$Feuille.Cells.Item($i, 50).EntireRow.Delete()	}
		Else {	Continue	}
	}
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
If ($args[1]) { $FicLog     = $RepLog + $ScriptName + "_" + $dat + "ONE_SHOT.log" } Else { $FicLog     = $RepLog + $ScriptName + "_REF.log" }
$LineSep    = "=" * 70

### Si le fichier LOG n'existe pas on le crée à vide
$Line = ">> DEBUT script de contrôle ESXi <<"
If (!(Test-Path $FicLog)) {
	$Line = (Get-Date -format G) + " " + $Line
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Line }
Else {	LogTrace ($Line)	}

### Test du contenu des paramètres de la ligne de commandes
If ($args[1]) {
	$Mode = "Mode 'Par vCenter avec filtrage'"
	$Mode_Var = 2
	Write-Host "$Mode [ID $Mode_Var]" -ForegroundColor Red -BackgroundColor White
	
	### Test de l'existence du fichier Excel de référence
	If (Test-Path ($RepLog + $ScriptName + "_REF.xlsx")) {	Write-Host "INFO: Le fichier Excel de référence '$($RepLog + $ScriptName + "_REF.xlsx")' est disponible. Comparaison possible..." -ForegroundColor Green	}
	Else {	Write-Host "INFO: Le fichier Excel de référence '$($RepLog + $ScriptName + "_REF.xlsx")' est indisponible. Comparaison impossible..." -ForegroundColor Red	}
	
	Copy-Item ($PathScript + "ESX_Health_Modele.xlsx") ($RepLog + $ScriptName + "_" + $dat + "_ONE_SHOT.xlsx")
	
	$FicRes = $RepLog + $ScriptName + "_" + $dat + "_ONE_SHOT.xlsx"
	$FicRes_Ref	= $RepLog + $ScriptName + "_REF.xlsx"
	
	### Définition du fichier Excel
	Add-Type -AssemblyName Microsoft.Office.Interop.Excel
	$Excel = New-Object -ComObject Excel.Application
	$ExcelWorkBook		= $Excel.WorkBooks.Open($FicRes)		# Ouverture du fichier *_[ONE SHOT].xlsx
	$ExcelWorkSheet		= $Excel.WorkSheets.item(1)				# Définition de la feuille Excel par défaut du fichier *_[ONE SHOT].xlsx
	$Excel.WorkSheets.item(2).Delete()							# Suppression de la 2ème feuille (inutile) du fichier *_[ONE SHOT].xlsx
	
	$ExcelWorkBook_Ref	= $Excel.WorkBooks.Open($FicRes_Ref)	# Ouverture du fichier *_REF.xlsx
	$ExcelWorkSheet_Ref	= $Excel.WorkSheets.item(2)				# Définition de la feuille Excel par défaut du fichier *_REF.xlsx
	
	$ExcelWorkSheet.Activate()
	$Excel.Visible		= $False
	
	$Global:vCenters	= $args[0] # Nom du vCenter
	$clustexc 			= $args[1] # clusters exclusq
	$esxexc				= $args[2] # Esx exclus
	$clustinc			= $args[3] # clusters inclus
	$esxinc				= $args[4] # Esx inclus

	$tabclustexc = @(); $tabesxexc = @(); $tabclustinc = @(); $tabesxinc = @()
	$clustexc = $clustexc.ToUpper().Trim()
	$esxexc   = $esxexc.ToUpper().Trim()
	$clustinc = $clustinc.ToUpper().Trim()
	$esxinc   = $esxinc.ToUpper().Trim()

	If ($clustexc -eq "" -or $clustexc -eq "NONE")	{	$clustexc = "AUCUN"	}
	If ($esxexc -eq "" -or $esxexc -eq "NONE")		{	$esxexc = "AUCUN"	}
	If ($clustexc -ne "AUCUN" )						{	$tabclustexc = $clustexc.split(",")	}
	If ($esxexc -ne "AUCUN" )						{	$tabesxexc = $esxexc.split(",")	}
	If ($clustinc -eq "" -or $clustinc -eq "NONE")	{	$clustinc = "TOUS"	}
	If ($esxinc -eq "" -or $esxinc -eq "NONE")		{	$esxinc = "TOUS"	}
	If ($clustinc -ne "TOUS" )						{	$tabclustinc = $clustinc.split(","); If ($clustinc.Contains(",")) {	$tabclustinc_Counter = ($clustinc.split(",").GetUpperBound(0) + 1)	} 	Else {	$tabclustinc_Counter = 1	}	}
	If ($esxinc -ne "TOUS" )						{	$tabesxinc = $esxinc.split(","); If ($esxinc.Contains(",")) {	$tabesxinc_Counter = ($esxinc.split(",").GetUpperBound(0) + 1)	} 			Else {	$tabesxinc_Counter = 1	}	}

	LogTrace ("Mode 'Par vCenter avec filtrage'")
	LogTrace ("Fichier liste des vCenter à traiter .. : $vCenters")
	Logtrace ("Cluster à exclure .................... : $clustexc")
	Logtrace ("ESX à exclure ........................ : $esxexc")
	Logtrace ("Cluster à prendre en compte .......... : $clustinc")
	Logtrace ("ESX à prendre en compte .............. : $esxinc")
	LogTrace ($LineSep + $vbcrlf)
	$TabVcc = $vCenters.split(",")
} Else {
	$Mode = "Mode 'Par vCenter sans aucun filtrage'"
	$Mode_Var = 1
	Write-Host "$Mode [ID $Mode_Var]" -ForegroundColor Red -BackgroundColor White
	
	### Test de l'existence du fichier Excel de référence
	If (Test-Path -Path $($RepLog + $ScriptName + "_REF.xlsx")) {
		$Reponse = [System.Windows.Forms.MessageBox]::Show("Attention: le fichier '$($RepLog + $ScriptName + "_REF.xlsx")' existe déjà, voulez-vous le remplacer ? ", "Confirmation" , 4, 32)
		If (($Reponse -eq "Yes") -or ($Reponse -eq "Oui")) {
			Move-Item ($RepLog + $ScriptName + "_REF.xlsx") ($RepLog + $ScriptName + "_REF_(BCKP_"+ $dat + ").xlsx")			# Sauvegarde du précédent fichier _REF.xlsx
			Copy-Item ($PathScript + "ESX_Health_Modele.xlsx") -Destination ($RepLog + $ScriptName + "_REF.xlsx")		# Copie du modèle pour modification par le script
			Write-Host "Le fichier '$($RepLog + $ScriptName + "_REF.xlsx")' a été sauvegardé puis écrasé" -ForegroundColor Red
			LogTrace ("Le fichier '$($RepLog + $ScriptName + "_REF.xlsx")' a été sauvegardé puis écrasé")	}
		Else	{
			Write-Host "Le fichier '$($RepLog + $ScriptName + "_REF.xlsx")' n'a pas été écrasé" -ForegroundColor Red
			LogTrace ("Le fichier '$($RepLog + $ScriptName + "_REF.xlsx")' n'a pas été écrasé")
			LogTrace ("FIN du script")
			Write-Host -NoNewLine "FIN du script. Appuyer sur une touche pour quitter...`r`n"
			$NULL = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
			Exit	}
	}
	Else {
		Copy-Item ($PathScript + "ESX_Health_Modele.xlsx") -Destination ($RepLog + $ScriptName + "_REF.xlsx")		# Copie du modèle pour modification par le script
	}
	
	$FicRes = $RepLog + $ScriptName + "_REF.xlsx"
	
	### Définition du fichier Excel
	$Excel = New-Object -ComObject Excel.Application
	$ExcelWorkBook	= $Excel.WorkBooks.Open($FicRes)
	$ExcelWorkSheet = $Excel.WorkSheets.item(1)
	$ExcelWorkSheet_Ref = $Excel.WorkSheets.item(2)
	$ExcelWorkSheet.Activate()
	$Excel.Visible	= $False
	Clear_Excel_Feuille -Feuille $ExcelWorkSheet
	
	$Global:vCenters = $args[0]		# Nom du vCenter
	LogTrace ("Mode 'Par vCenter sans aucun filtrage'")
	LogTrace ("Fichier liste des vCenter à traiter .. : $vCenters")
	LogTrace ($LineSep + $vbcrlf)
	$TabVcc = $vCenters.split(",")	}

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
	If ($rccnx -ne $NULL) {	If ($rccnx.Isconnected) { $topCnxVcc = "1" } }

	If ($topCnxVcc -ne "1") { LogTrace ("ERREUR: Connexion KO au vCenter $vCenter - Arrêt du script")
		Write-Host "ERREUR: Connexion KO au vCenter $vCenter - Arrêt du script" -ForegroundColor White -BackgroundColor Red	
		$rc += 1
		Exit $rc }
	Else { LogTrace ("SUCCES: Connexion OK au vCenter $vCenter" + $vbcrlf)
		Write-Host "SUCCES: Connexion OK au vCenter $vCenter" -ForegroundColor Black -BackgroundColor Green	}

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
			If (($Mode_Var -eq 2) -and ($clustinc -ne "TOUS"))	{
				$ClusterNom = $Cluster.Name
				$Global:Cluster_Counter = $tabclustinc_Counter
				If ($tabclustexc -Contains $ClusterNom) {
					Logtrace (" * Exclusion du cluster '$ClusterNom'...")
					Write-Host " * Exclusion " -NoNewLine -ForegroundColor Red
					Write-Host "du cluster " -NoNewLine
					Write-Host "'$ClusterNom'..." -ForegroundColor DarkYellow
					Continue	}
				If (($tabclustinc.length -ne 0) -and ($tabclustinc -notContains $ClusterNom)) {
					Logtrace (" * Cluster '$ClusterNom' absent des clusters à traiter...")
					Write-Host " * Cluster " -NoNewLine
					Write-Host "'$ClusterNom' " -NoNewLine -ForegroundColor DarkYellow
					Write-Host "absent " -NoNewLine -ForegroundColor Red
					Write-Host "des clusters à traiter..."
					Continue	}
			}

			$noCluster += 1; $Cluster_Inc += 1
			$Global:ExcelLine_Start += $ESX_Counter

			LogTrace ($vbcrlf + "Traitement CLUSTER $Cluster n°$noCluster sur $Cluster_Counter {$Cluster_Inc}")
			Write-Host "Traitement CLUSTER [#$noCluster/$Cluster_Counter] " -NoNewLine
			Write-Host "{$Cluster_Inc} " -ForegroundColor Red -NoNewLine
			Write-Host "$Cluster... ".ToUpper() -ForegroundColor Yellow -NoNewLine
			Write-Host "En cours" -ForegroundColor Green
			
			$Global:oESX = Get-vmHost -Location $Cluster
			$Global:ESX_Counter = $oESX.Count
			$Global:no_ESX = 0
			
			### Boucle de traitement des ESXi composants le cluster dont ceux contenus dans les paramètres de la ligne de commandes
			ForEach($ESX in Get-vmHost -Location $Cluster) {
				If (($Mode_Var -eq 2) -and ($esxinc -ne "TOUS")) {
					$Global:ESX_Counter = $tabesxinc_Counter
					If ($tabesxexc -Contains $ESX) {
						Logtrace (" * Exclusion de l'ESXi '$ESX'...")
						Write-Host " * Exclusion " -NoNewLine -ForegroundColor Red
						Write-Host "de l'ESXi " -NoNewLine
						Write-Host "'$ESX'..." -ForegroundColor DarkYellow
						Continue	}
					If (($tabesxinc.length -ne 0) -and ($tabesxinc -notContains $ESX)) {
						Logtrace (" * ESXi '$ESX' absent des ESX à traiter...")
						Write-Host " * ESXi " -NoNewLine
						Write-Host "'$ESX' " -NoNewLine -ForegroundColor DarkYellow
						Write-Host "absent " -NoNewLine -ForegroundColor Red
						Write-Host "des ESXi à traiter..."
						Continue	}
				}
 
				$no_ESX += 1; $no_inc_ESX += 1
				
				$StartTime = Get-Date -Format HH:mm:ss
				LogTrace ("Traitement ESXi '$ESX' n°$no_ESX sur $($ESX_Counter) {$no_inc_ESX}")
				Write-Host "[$StartTime] Traitement ESXi [#$no_ESX/$($ESX_Counter)] " -NoNewLine
				Write-Host "{$no_inc_ESX} " -ForegroundColor Red -NoNewLine
				Write-Host "'$ESX'... ".ToUpper() -ForegroundColor Yellow -NoNewLine
				Write-Host "En cours" -ForegroundColor Green

				If ($ESX.PowerState -ne "PoweredOn") {
					For ($i = 2; $i -le 37; $i++) {	$ExcelWorkSheet.Cells.Item($no_ESX + 1, $i) = "NA"	}
						$ExcelWorkSheet.Cells.Item($no_inc_ESX + 1, 13) = "PoweredOff" # Ligne, Colonne
					Continue }
				
				### Exécution des fonctions de récupération des données ESX
				Get-ESX_HARD -vmHost $ESX		# Récupération matérielle
				Get-ESX_CONFIG -vmHost $ESX		# Récupération de la configuration matérielle
				Get-ESX_SAN -vmHost $ESX		# Récupération des données stockage SAN
				Get-ESX_HOST -vmHost $ESX		# Récupération de la configuration ESXi
				Get-ESX_NETWORK -vmHost $ESX	# Récupération de la configuration réseaux
				Get-HP_ILO -vmHost $ESX			# Récupération de la configuration iLO
				Get-ESX_Compliant				# Analyse de la conformité standard
								
				$EndTime = Get-Date -Format HH:mm:ss
				LogTrace ("MISE A JOUR des données pour l'ESX '$ESX'" + $vbcrlf)
				Write-Host "[$EndTime] Mise à jour des données Excel " -NoNewLine
				Write-Host "Zone [L$ExcelLine_Start..$($ExcelLine_Start + $ESX_Counter - 1) *L$($no_inc_ESX + 1)*]`r`n"  -ForegroundColor Yellow
				
				$ExcelWorkBook.Save()

				### Sélection de la fonction selon le mode de départ
				If ($Mode_Var -eq 1) {	If ($no_ESX -eq $($oESX.Count)) {	Get-ESX_Compare_Full }	}	# Dans le cas d'une découverte complète du périmètre
				If (($Mode_Var -eq 2) -And (!(Test-Path $($RepLog + $ScriptName + "_REF.xlsx")))) {	If ($no_ESX -eq $ESX_Counter) {	Get-ESX_Compare_Part_Ref; 	Break }	} # Dans le cas d'une vérification de conformité unitaire si le fichier Excel de référence existe
				If (($Mode_Var -eq 2) -And (Test-Path $($RepLog + $ScriptName + "_REF.xlsx"))) 	{	If ($no_ESX -eq $ESX_Counter) {	Get-ESX_Compare_Part; 		Break }	} # Dans le cas d'une vérification de conformité unitaire si le fichier Excel de référence n'existe pas
			}
		}
	}
	
	### Exécution des fonctions de récupération des valeurs ESX
	If ($Mode_Var -eq 1) {	Get-Cluster_Compare	}
	
	LogTrace ("DECONNEXION et FIN du traitement depuis le VCENTER $vCenter`r`n")
	Disconnect-VIServer -Server $vCenter -Force -Confirm:$False
}

#$Excel.ActiveWorkBook.SaveAs($FicRes); Write-Host "Enregistrement du fichier Excel [Terminé]"; LogTrace ("Enregistrement du fichier Excel [Terminé]")
$Excel.WorkBooks.Close($True); Write-Host "Fermeture du classeur Excel [Terminé]"; LogTrace ("Fermeture du classeur Excel [Terminé]")
$Excel.Quit(); Write-Host "Fermeture du programme Excel [Terminé]"; LogTrace ("Fermeture du programme Excel [Terminé]")

Send_Mail; Write-Host "Envoi du mail avec les fichiers LOG et XLSX"; LogTrace ("Envoi du mail avec les fichiers LOG et XLSX")
Write-Host -NoNewLine "FIN du script. Appuyer sur une touche pour quitter...`r`n"
$NULL = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")