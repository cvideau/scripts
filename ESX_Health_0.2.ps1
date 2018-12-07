#[01] Nom ESXi
#[02] HARD - Constructeur
#[03] HARD - Modèle
#[04] HARD - N° de Série
#[05] HARD - Version BIOS
#[06] HARD - Date du BIOS
#[07] HARD - Etat Carte SD
#[08] CONF - Datacenter
#[09] CONF - Cluster Parent
#[10] CONF - Etat ESX
#[11] CONF - vCenter Parent
#[12] CONF - Balise vSphere
#[13] CONF - Statut ESX
#[14] CONF - Version
#[15] CONF - Serveurs NTP
#[16] CONF - Zone de temps
#[17] CONF - Heure locale
#[18] CONF - Niveau de conformité
#[19] CONF - Etat du démon NTP
#[20] SAN - Total des Chemins
#[21] SAN - Total des chemin 'MORTS'
#[22] SAN - Total des chemin 'ACTIFS'
#[23] SAN - Etat de la redondance des chemins
#[24] TOOLS - Present
#[25] TOOLS - Version
#[26] HOST - Current CPU Policy
#[27] HOST - Mode HA
#[28] HOST - Hyperthreading
#[29] HOST - EVC
#[30] HOST - Etat des alarmes
#[31] HOST - TPS Salting
#[32] HOST - Large Pages RAM
#[33] NETWORK - VLAN
#[34] NETWORK - Adaptateurs
#[35] NETWORK - vMotion IP


Function Get-ESX_HARD { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs HARDWARE..."
	$ESXCLI2 = Get-ESXCLI -VMHost $ESX.Name -Server $vCenter
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 1) = $ESX.Name
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 2) = $ESXCLI2.Hardware.Platform.Get.Invoke().VendorName
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 3) = $ESXCLI2.Hardware.Platform.Get.Invoke().ProductName
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 4) = $ESXCLI2.Hardware.Platform.Get.Invoke().SerialNumber
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 5) = $ESX.ExtensionData.Hardware.BiosInfo.BiosVersion
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 6) = $ESX.ExtensionData.Hardware.BiosInfo.ReleaseDate
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 7) = $ESXCLI2.Storage.Core.Path.List.Invoke().AdapterTransportDetails | Where { $_.Device -eq "mpx.vmhba32:C0:T0:L0" }
}

	
Function Get-ESX_CONFIG { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs CONFIGURATION..."
	$ESXCLI2 = Get-ESXCLI -VMHost $ESX.Name -Server $vCenter
	$Global:oESXTAG = Get-TagAssignment -Entity $ESX
	$Global:oESX_NTP = Get-VMHostNtpServer -VMHost $ESX.Name -Server $vCenter
	
	If ((Get-VmHostService -VMHost $ESX -Server $vCenter | Where-Object {$_.key -eq "ntpd"}).Running -eq "True") {	$ntpd_status = "Running"	} Else { $ntpd_status = "Not Running" }
	If ($ESXCLI2.System.MaintenanceMode.Get.Invoke() -eq "Enabled") { $State = "Maintenance" } Else { $State = "Connected"}
	If ((Get-Compliance -Entity $ESX -Detailed | WHERE {$_.NotCompliantPatches -ne $NULL} | SELECT Status).Count -gt 0) { $ComplianceLevel = "Baseline - Not compliant" } Else { $ComplianceLevel = "Baseline - Compliant" }
	$Global:UTC = Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}
	$Global:UTC_Var1 = $UTC.Config.DateTimeInfo.TimeZone.Name
	$Global:UTC_Var2 = $UTC.Config.DateTimeInfo.TimeZone.GmtOffset
	$Global:Gateway = Get-VmHostNetwork -Host $ESX  -Server $vCenter | Select VMkernelGateway -ExpandProperty VirtualNic | Where {$_.VMotionEnabled} | Select -ExpandProperty VMkernelGateway
	
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 8) = "$DC" # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 9) = "$Cluster" # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 10) = $State # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 11) = $vCenter # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 12) = "$oESXTAG" # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 13) = "PoweredOn" # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 14) = $ESX.Version + " - Build " + $ESX.Build
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 15) = $oESX_NTP
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 16) = $UTC_Var1 + "+" + $UTC_Var2 # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 17) = $ESXCLI2.Hardware.Clock.Get.Invoke() # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 18) = $ComplianceLevel # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 19) = $NTPD_status
	
	If (($ESX.Build -eq "7611317") -and ($oESXTAG -ne $NULL)) {$Compliance_Build = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 14).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 14).Font.Bold = $False} Else {$Compliance_Build = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 14).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 14).Font.Bold = $True}
	If (($oESX_NTP -eq $Gateway) -and ($oESXTAG -ne $NULL)) {$Compliance_NTP = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 15).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 15).Font.Bold = $False} Else {$Compliance_NTP = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 15).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 15).Font.Bold = $True}
	If (($UTC_Var2 -eq "0") -and ($oESXTAG -ne $NULL)) {$Compliance_TimeOffset = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 16).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 16).Font.Bold = $False} Else {$Compliance_TimeOffset = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 16).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 16).Font.Bold = $True}
	If (($ComplianceLevel -eq "Baseline - Compliant") -and ($oESXTAG -ne $NULL)) {$Compliance_Level = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 18).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 18).Font.Bold = $False} Else {$Compliance_Level = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 18).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 18).Font.Bold = $True}
	If (($NTPD_status -eq "Running") -and ($oESXTAG -ne $NULL)) {$Compliance_NTPd = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 19).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 19).Font.Bold = $False} Else {$Compliance_NTPd = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 19).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 19).Font.Bold = $True}
}
	
	
Function Get-ESX_SAN { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs SAN..."
	Get-VMHostStorage -RescanAllHba -VMHost $ESX -Server $vCenter | Out-Null
	$ESXCLI2 = Get-ESXCLI -VMHost $ESX -Server $vCenter
	$Global:PathsSum = ($ESXCLI2.storage.core.path.list.invoke() | Where {$_.Plugin -eq 'PowerPath'} | Select Device -Unique).Count
	$Global:PathsDead = ($ESXCLI2.storage.core.path.list.invoke() | Where {($_.State -eq "dead") -and ($_.Plugin -eq 'PowerPath')} | Select Device -Unique).Count
	$Global:PathsActive = ($ESXCLI2.storage.core.path.list.invoke() | Where {($_.State -eq "active") -and ($_.Plugin -eq 'PowerPath')} | Select Device -Unique).Count
	$Global:HBAAdapter = (Get-VMHostHba -VMHost $ESX -Type "FibreChannel").Count
	
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 20) = "$PathsSum" + " (" + $HBAAdapter + " adapters)" # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 21) = $PathsDead # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 22) = $PathsActive # Ligne, Colonne
	If ($PathsActive -eq $PathsSum) { $PathsRedondance = "OK" } Else { $PathsRedondance = "NOK" }
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 23) = $PathsRedondance # Ligne, Colonne
	
	If (($PathsRedondance -eq "OK") -and ($oESXTAG -ne $NULL)) {$Compliance_PathsDead = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 23).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 23).Font.Bold = $False} Else {$Compliance_PathsDead = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 23).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 23).Font.Bold = $True}
}


Function Get-ESX_TOOLS { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs VMTOOLS..."
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 24) = $vmToolsPackage # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 25) = $VersionInstalled # Ligne, Colonne
}


Function Get-ESX_HOST { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs HOTE..."
	$Global:CurrentEVCMode = (Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Summary.CurrentEVCModeKey
	$Global:MaxEVCMode = (Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Summary.MaxEVCModeKey
	$Global:CPUPerformance = (Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Hardware.CpuPowerManagementInfo.CurrentPolicy
	$Global:HyperThreading = (Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Config.HyperThread.Active
	If ((Get-VmHostService -VMHost $ESX -Server $vCenter | Where-Object {$_.key -eq "vmware-fdm"}).Running -eq "True") {	$HAEnabled = "Running"	} Else { $HAEnabled = "Not Running" }
	$Global:Alarm = (Get-VMHost -Name $ESX).ExtensionData.AlarmActionsEnabled
	If ($Alarm -eq "True") { $Alarm = "Enabled" } Else { $Alarm = "Disabled" }
	If ((Get-AdvancedSetting -Entity $ESX -Name "Mem.AllocGuestLargePage").Value -eq "0") { $LPages = "Disabled"	} Else { $LPages = "Enabled" }
	If ((Get-AdvancedSetting -Entity $ESX -Name "Mem.ShareForceSalting").Value -eq "0") { $TPS = "Enabled" } Else { $TPS = "Disabled" }

	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 26) = $CPUPerformance # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 27) = $HAEnabled # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 28) = $HyperThreading # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 29) = "Current: " + $CurrentEVCMode + ", Max: " + $MaxEVCMode # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 30) = $Alarm # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 31) = $TPS # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 32) = $LPages # Ligne, Colonne
	
	If (($CPUPerformance -eq "High Performance") -and ($oESXTAG -ne $NULL)) {$Compliance_CPUPolicy = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 26).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 26).Font.Bold = $False} Else {$Compliance_CPUPolicy = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 26).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 26).Font.Bold = $True}
	If (($HAEnabled -eq "Running") -and ($oESXTAG -ne $NULL)) {$Compliance_HA = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 27).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 27).Font.Bold = $False} Else {$Compliance_HA = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 27).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 27).Font.Bold = $True}
	If (($HyperThreading -eq "VRAI") -and ($oESXTAG -ne $NULL)) {$Compliance_HT = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 28).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 28).Font.Bold = $False} Else {$Compliance_HT = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 28).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 28).Font.Bold = $True}
	If (($Alarm -eq "Enabled") -and ($oESXTAG -ne $NULL)) {$Compliance_Alarm = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 30).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 30).Font.Bold = $False} Else {$Compliance_Alarm = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 30).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 30).Font.Bold = $True}
	If (($TPS -eq "Enabled") -and ($oESXTAG -ne $NULL)) {$Compliance_TPS = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 31).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 31).Font.Bold = $False} Else {$Compliance_TPS = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 31).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 31).Font.Bold = $True}
	If (($LPages -eq "Enabled") -and ($oESXTAG -ne $NULL)) {$Compliance_LPages = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 32).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 32).Font.Bold = $False} Else {$Compliance_LPages = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 32).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 32).Font.Bold = $True}
}


Function Get-ESX_NETWORK { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	Write-Host " * Récupération des valeurs RESEAUX..."
	$ESXCLI2 = Get-ESXCLI -VMHost $ESX -Server $vCenter
	$ESXvLAN = Get-VirtualPortGroup -VMHost $ESX -Server $vCenter -Distributed | Select Name
	ForEach ($DPortgroup in $ESXvLAN)	{
		$vLANID = ($ESXvLAN.Name).SubString(0, 4)
		Get-Unique -AsString -InputObject $vLANID | Out-Null
	}
	$vLANID = $vLANID -Replace "0000","" -Replace "_", "" -Replace "[^0-9]","" -Replace "  ",""
	$vLANID = "(" + $vLANID.Count + ")`n " + $vLANID
	
	$Global:ESXAdapter = $ESXCLI2.Network.nic.list.Invoke() | Where {$_.Link -eq "Up"}
	$Global:ESXAdapterName = $ESXCLI2.Network.nic.list.Invoke() | Where {$_.Link -eq "Up"} | Select Name, Speed, Duplex | Out-String
	$Global:ESXAdapterName = $ESXAdapterName -Replace "-","" -Replace "Name","" -Replace "Speed","" -Replace "Duplex","" -Replace "`r`n","" -Replace " ","" -Replace "10000", " 10000 " -Replace "vm", "`r`nvm"
	$Global:ESXAdapterCount = ($ESXCLI2.Network.nic.list.Invoke() | Where {$_.Link -eq "Up"}).Count
	
	$vMotionIP = Get-VMHostNetworkAdapter -VMHost $ESX | Where {$_.DeviceName -eq "vmk1"} | Select IP | Out-String
	$vMotionIP = $vMotionIP -Replace "-","" -Replace "IP","" -Replace "`n","" -Replace " ",""
	$vMotionEnabled = (Get-View -ViewType HostSystem -Filter @{"Name" = $ESX.Name}).Summary.Config.VmotionEnabled
	If ($vMotionEnabled -eq $True) {$vMotionEnabled = "vMotion enabled"} Else {$vMotionEnabled = "vMotion disabled"}

	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 33) = $vLANID # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 34) = "(" + $ESXAdapterCount + ") " + $ESXAdapterName # Ligne, Colonne
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 35) = $vMotionIP + " (" + $vMotionEnabled + ")" # Ligne, Colonne
	
	If ((($ESXAdapterName -notcontains "Half") -or ($ESXAdapterName -notcontains "1000")) -and ($oESXTAG -ne $NULL)) {$Compliance_NetworkFlow = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 34).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 34).Font.Bold = $False} Else {$Compliance_Level = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 34).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 34).Font.Bold = $True}
	If (($vMotionEnabled -eq "vMotion enabled") -and ($oESXTAG -ne $NULL)) {$Compliance_vMotion = "Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 35).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 35).Font.Bold = $False} Else {$Compliance_vMotion = "Not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 35).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 35).Font.Bold = $True}
}

Function Get-HP_ILO { Param( [Parameter(Position = 0, ValueFromPipeline=$True, Mandatory=$True)] [String]$vmHost )
	$ExcelWorkSheet.Cells.Item($no_ESX + 1, 37) = $iLOversion # Ligne, Colonne
}

Function Get-ESX_Compliant {
	If (($Compliance_Build -eq "Compliant") `
		-and ($Compliance_NTP -eq "Compliant") `
		-and ($Compliance_TimeOffset -eq "Compliant") `
		-and ($Compliance_Level -eq "Compliant") `
		-and ($Compliance_NTPd -eq "Compliant") `
		-and ($Compliance_PathsDead -eq "Compliant") `
		-and ($Compliance_CPUPolicy -eq "Compliant") `
		-and ($Compliance_HA -eq "Compliant") `
		-and ($Compliance_HT -eq "Compliant") `
		-and ($Compliance_Alarm -eq "Compliant") `
		-and ($Compliance_TPS -eq "Compliant") `
		-and ($Compliance_LPages -eq "Compliant") `
		-and ($Compliance_NetworkFlow -eq "Compliant") `
		-and ($Compliance_vMotion -eq "Compliant")) { $ExcelWorkSheet.Cells.Item($no_ESX + 1, 38) = "ESX Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 38).Font.ColorIndex = 50; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 38).Font.Bold = $True }
	Else { $ExcelWorkSheet.Cells.Item($no_ESX + 1, 38) = "ESX not Compliant"; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 38).Font.ColorIndex = 3; $ExcelWorkSheet.Cells.Item($no_ESX + 1, 38).Font.Bold = $True }
}

Function Get-ESX_Compare {
	$TotalPaths = @($ExcelLine_Start .. $($ExcelLine_Start + $ESX_Counter))
	$TotalPaths | ForEach {
		$ExcelWorkSheet.Cells.Item($ExcelLine_Start, 20).Text.Substring(0, $ExcelWorkSheet.Cells.Item($ExcelLine_Start, 20).Text.IndexOf(" "))
	}
	Write-Host $TotalPaths
}

Import-Module "VMware.VimAutomation.Core"
Add-PsSnapin VMware.VumAutomation
Clear-Host

$Global:vCenters = $args[0] # Nom du vCenter
$clustexc = $args[1] # clusters exclusq
$esxexc   = $args[2] # Esx exclus
$clustinc = $args[3] # clusters inclus
$esxinc   = $args[4] # Esx inclus
$Global:ExcelLine_Start = 1

# Bouchon pour tests
#$vcenters = "swmuzv1vcszd.zres.ztech"
#$clustexc = "AUCUN"
#$esxexc = "AUCUN"
#$clustinc = "CL_MU_HDI_Z80,CL_MU_HDM_Z80"
#$esxinc = "sxmuzhvhdich.zres.ztech,sxmuzhvhdidk.zres.ztech"
#

$PathScript = ($myinvocation.MyCommand.Definition | Split-Path -parent) + "\"
$user       = "CSP_SCRIPT_ADM"
$fickey     = "D:\Scripts\Credentials\key.crd"
$ficcred    = "D:\Scripts\Credentials\vmware_adm.crd"
$key        = Get-content $fickey
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
$FicRes     = $PathScript + "ESX_Health.xlsx"
$LineSep    = "=" * 70

### Si le fichier LOG n'existe pas on le crée à vide
$Line = ">> DEBUT script de contrôle ESXi <<"
If (!(Test-Path $FicLog)) {
	$Line = (Get-Date -format G) + " " + $Line
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Line
} Else {
	LogTrace ($Line)
}

$tabclustexc = @() ; $tabesxexc = @() ; $tabclustinc = @() ; $tabesxinc = @()

$clustexc = $clustexc.ToUpper().Trim()
$esxexc   = $esxexc.ToUpper().Trim()
$clustinc = $clustinc.ToUpper().Trim()
$esxinc   = $esxinc.ToUpper().Trim()

If ($clustexc -eq "" -or $clustexc -eq "NONE") {
  $clustexc = "AUCUN"
}

If ($esxexc -eq "" -or $esxexc -eq "NONE") {
  $esxexc = "AUCUN"
}

If ($clustexc -ne "AUCUN" ) {
  $tabclustexc = $clustexc.split(",")
}

If ($esxexc -ne "AUCUN" ) {
  $tabesxexc = $esxexc.split(",")
}

If ($clustinc -eq "" -or $clustinc -eq "NONE") {
  $clustinc = "TOUS"
}

If ($esxinc -eq "" -or $esxinc -eq "NONE") {
  $esxinc = "TOUS"
}

If ($clustinc -ne "TOUS" ) {
  $tabclustinc = $clustinc.split(",")
}

If ($esxinc -ne "TOUS" ) {
  $tabesxinc = $esxinc.split(",")
}

LogTrace ("Fichier liste des vCenter à traiter .. : $vCenters")
Logtrace ("Cluster à exclure .................... : $clustexc")
Logtrace ("ESX à exclure ........................ : $esxexc")
Logtrace ("Cluster à prendre en compte .......... : $clustinc")
Logtrace ("ESX à prendre en compte .............. : $esxinc")
LogTrace ($LineSep + $vbcrlf)
$TabVcc = $vCenters.split(",")

### Définition des entêtes du fichier de sortie
$Excel = New-Object -ComObject Excel.Application
$ExcelWorkBook  = $Excel.WorkBooks.Open($FicRes)
$ExcelWorkSheet = $Excel.WorkSheets.item(1)
$ExcelWorkSheet.Activate()
$Excel.Visible = $False

# Modif PCO
$Global:no_ESX = 0

ForEach ($vCenter in $TabVcc) {	LogTrace ("DEBUT du traitement du vCenter $vCenter")
	Write-Host "DEBUT du traitement du vCenter " -NoNewLine
	Write-Host "$vCenter... ".ToUpper() -ForegroundColor Yellow -NoNewLine
	Write-Host "En cours" -ForegroundColor Green	
	
	$rccnx = Connect-VIServer -Server $vcenter -Protocol https -Credential $Credential
	
	$topCnxVcc = "0"
	If ($rccnx -ne $null) {	If ($rccnx.Isconnected) { $topCnxVcc = "1" } }

	If ($topCnxVcc -ne "1") { LogTrace ("ERREUR: Connexion KO au vCenter $vCenter => Arrêt du script")
		Write-Host "ERREUR: Connexion KO au vCenter $vCenter => Arrêt du script" -ForegroundColor White -BackgroundColor Red
		$rc += 1
		Exit $rc }
	Else { LogTrace ("SUCCES: Connexion OK au vCenter $vCenter" + $vbcrlf)
		Write-Host "SUCCES: Connexion OK au vCenter $vCenter" -ForegroundColor Black -BackgroundColor Green	}


	$Global:noDatacenter = 0
	$Global:oDatacenters = Get-Datacenter | Sort Name
	$Global:Datacenter_Counter = $oDatacenters.Count
	ForEach($DC in $oDatacenters){ $noDatacenter += 1
		LogTrace ("Traitement du DATACENTER $DC n°$noDatacenter sur $Datacenter_Counter" + $vbcrlf)
		Write-Host "Traitement du DATACENTER [#$noDatacenter/$Datacenter_Counter] " -NoNewLine
		Write-Host "$DC... ".ToUpper() -ForegroundColor Yellow -NoNewLine
		Write-Host "En cours" -ForegroundColor Green	

		$Global:noCluster = 0
		$Global:oClusters = Get-Cluster -Location $DC | Sort Name
		$Global:Cluster_Counter = $oClusters.Count

#		ForEach($Cluster in Get-Cluster -Location $DC){
		ForEach($Cluster in Get-Cluster -Location $DC | Sort Name){

# Modif PCO
			$clustnom = $Cluster.Name
  If ($tabclustexc -contains $clustnom) {
     Logtrace ("Exclusion du cluster $clustnom => BYPASS du cluster")
     Continue
   }

   If ($tabclustinc.length -ne 0 -and $tabclustinc -notcontains $clustnom) {
     Logtrace ("Cluster $clustnom absent des clusters à prendre en compte => BYPASS")
     Continue
   }
# Fin Modif

			$noCluster += 1

			LogTrace ("Traitement du CLUSTER $Cluster n°$noCluster sur $Cluster_Counter")
			Write-Host "Traitement du CLUSTER [#$noCluster/$Cluster_Counter] " -NoNewLine
			Write-Host "$Cluster... ".ToUpper() -ForegroundColor Yellow -NoNewLine
			Write-Host "En cours" -ForegroundColor Green
			
			$Global:oESX = Get-vmHost -Location $Cluster | Sort Name
			$Global:ESX_Counter = $oESX.Count
			
			ForEach($ESX in Get-VMHost -Location $Cluster) {

# Modif PCO
    $esxnom = $ESX.Name
     
    If ($tabesxexc -contains $esxnom) {
      Logtrace ("Exclusion de l'ESX $esxnom => BYPASS de l'ESX")
      Continue
    }

    If ($tabesxinc.length -ne 0 -and $tabesxinc -notcontains $esxnom) {
      Logtrace ("ESX $esxnom absent des ESX à prendre en compte => BYPASS")
      Continue
    }
 # Fin Modif
 
				$no_ESX += 1
				$ExcelLine_Start += 1
				
				LogTrace ("Traitement de l'ESX $ESX n°$no_ESX sur $ESX_Counter")
				Write-Host "Traitement de l'ESX [#$no_ESX/$ESX_Counter] " -NoNewLine
				Write-Host "$ESX... ".ToUpper() -ForegroundColor Yellow -NoNewLine
				Write-Host "En cours" -ForegroundColor Green

				If ($ESX.PowerState -ne "PoweredOn") {
					For ($i = 2; $i -le 36; $i++) {	$ExcelWorkSheet.Cells.Item($no_ESX + 1, $i) = "NA"	}
					$ExcelWorkSheet.Cells.Item($no_ESX + 1, 13) = "PoweredOff" # Ligne, Colonne
					Continue }
				
				Get-ESX_HARD -vmHost $ESX
				Get-ESX_CONFIG -vmHost $ESX
				Get-ESX_SAN -vmHost $ESX
				Get-ESX_TOOLS -vmHost $ESX
				Get-ESX_HOST -vmHost $ESX
				Get-ESX_NETWORK -vmHost $ESX
				Get-HP_ILO -vmHost $ESX
				Get-ESX_Compliant
				If ($no_ESX -eq $ESX_Counter) {	Get-ESX_Compare }
				
				LogTrace ("MISE A JOUR des informations dans le fichier de sortie pour l'ESX $ESX" + $vbcrlf)
				Write-Host "Mise à jour du fichier de sortie pour l'ESX " -NoNewLine
				Write-Host "$ESX... "  -ForegroundColor Yellow
				$ExcelWorkBook.Save()
			}
		}
	}
	LogTrace ("DECONNEXION et FIN du traitement depuis le vCenter $vCenter`r`n")
	Disconnect-VIServer -Server $vCenter –Force –Confirm:$False
}
$ExcelWorkBook.Save()
$ExcelWorkBook.Close()
$Excel.Quit()
LogTrace ("FIN du script")