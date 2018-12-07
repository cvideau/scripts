### ADVERTISEMENT
### (2018) Christophe VIDEAU - Version 1.0
Clear-Host

Write-Host "FOR A BETTER TIME EXECUTION, SORT 'ClusterList.csv' BY VCENTER !`r`nIF NEEDED, PLEASE QUIT THIS PROGRAM AND MODIFY 'ClusterList.csv'" -ForegroundColor White -BackgroundColor Red

### FUNCTION: Create and append LOG file
Function LogTrace ($Message){
	$Message = (Get-Date -format G) + " " + $Message
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Message -Append }

### Load assembly & context
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#$ErrorActionPreference= 'SilentlyContinue'
Add-PSsnapin VMware.VimAutomation.Core
### https://kb.vmware.com/selfservice/microsites/search.do?language=en_US&cmd=displayKC&externalId=2009857

### TEST CSV file exist
If (Test-Path $PSScriptRoot\ClusterList.csv) {
	$ClusterList = Import-CSV -Delimiter ';' $PSScriptRoot\ClusterList.csv
	$CountLine = (Get-ChildItem -Path $PSScriptRoot\ClusterList.csv -Recurse | Get-Content | Measure-Object -Line | Select -Expand Lines) -1 }
Else {
	### EXIT Program
	[System.Windows.Forms.MessageBox]::Show("CSV file ($PSScriptRoot\ClusterList.csv) is missing, program is aborted !", "Abort" ,0 , 16)
	Write-Host "Program is aborted !" -ForegroundColor White -BackgroundColor Red
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject "[ERROR] CSV file is missing, program aborted !" -Append
	$Host.Exit()
}

$Reponse = [System.Windows.Forms.MessageBox]::Show("Are you sure that you want to modify TPS parameters on each clusters presents in the input file ?", "Confirm" , 4, 32)
If ($Reponse -eq "Yes") {

	### USER identification initialization
	$PathScript = ($myinvocation.MyCommand.Definition | Split-Path -Parent) + "\"
	$user       = "CSP_SCRIPT_ADM"
	$fickey     = "D:\Scripts\Credentials\key.crd "
	$ficcred    = "D:\Scripts\Credentials\vmware_adm.crd"
	$key        = Get-Content $fickey
	$pwd        = Get-Content $ficcred | ConvertTo-SecureString -key $key
	$Credential = New-Object System.Management.Automation.PSCredential $user, $pwd

	### LOG file initialization
	$Timestamp	= Get-Date -UFormat "%Y_%m_%d_%H_%M_%S"
	$vbcrlf		= "`r`n"
	$ScriptName	= [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition)
	$dat		= Get-Date -UFormat "%Y_%m_%d_%H_%M_%S"

	$RepLog     = $PathScript + "LOG"
	If (!(Test-Path $RepLog)) { New-Item -Path $PathScript -ItemType Directory -Name "LOG"}
	$RepLog     = $RepLog + "\"
	$FicLog     = $RepLog + $ScriptName + "_" + $dat + ".log"
	$FicRes     = $RepLog + $ScriptName + "_" + $dat + "_SET_ESX_TPS.csv"
	$LineSep    = "=" * 70

	$Line = ">> BEGIN Script modification of TPS parameters <<"
	If (!(Test-Path $FicLog)) {
		$Line = (Get-Date -format G) + " " + $Line
		Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Line
	} Else {
		LogTrace ($Line)
	}
	Write-Host "LogFile initialization is done !" -ForegroundColor Green
	LogTrace ($LineSep + $vbcrlf)

	### DEFINE outfile column header
	$LOG_Head  = "ESXi;vCenter;Cluster;Action;Resultat 'LPAGE';Resultat 'SALTING';AllocGuestLargePage (Avant);ShareForceSalting (Avant);AllocGuestLargePage (Apres);ShareForceSalting (Apres)"
	Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Head

	ForEach ($Item in $ClusterList) {
		$Cluster					= $Item.("Cluster")				### CLUSTER
		$vCenter					= $Item.("vCenter")				### vCenter
		$AllocGuestLargePage		= $Item.("AllocGuestLargePage")	### RAM Large Page Parameter
		$ShareForceSalting			= $Item.("ShareForceSalting")	### RAM Share Force Salting Parameter
		$CountLine					-= 1							### DECREASE FILE LINE VARIABLE
		$Maintenance_TimeOut_Delay	= 600							### Timeout delay in secondes
		$Resultat_LPAGE				= "NOTHING WAS DONE"							### INIT Variable
		$Resultat_SALTING			= "NOTHING WAS DONE"							### INIT Variable
		#[System.Windows.Forms.MessageBox]::Show("$Cluster $Action")
		
		### CSV file CHECK
		If ([string]::IsNullOrEmpty($Cluster))	{
			LogTrace ("ERROR: Parameter 'cluster' (CSV file as input) is empty. Please enter the name of a target 'cluster' in the CSV file")
			Write-Host "[ERROR]: Parameter 'cluster' (CSV file as input) is empty. Please enter the name of a target 'cluster' in the CSV file"
			Continue	}

		If ([string]::IsNullOrEmpty($vCenter))	{
			LogTrace ("ERROR: Parameter 'vcenter' (CSV file as input) is empty. Please enter the name of a target 'cluster' in the CSV file")
			Write-Host "[ERROR]: Parameter 'vcenter' (CSV file as input) is empty. Please enter the name of a target 'cluster' in the CSV file"
			Continue	}
				
			
		If ([string]::IsNullOrEmpty($AllocGuestLargePage) -Or (-Not [string]($AllocGuestLargePage -as [int])))	{
			LogTrace ("ERROR: Parameter 'AllocGuestLargePage' (CSV file as input) is empty or incorrect. Please enter a value for this parameter in the CSV file")
			Write-Host "ERROR: Parameter 'AllocGuestLargePage' (CSV file as input) is empty or incorrect. Please enter a value for this parameter in the CSV file"
			Continue	}
			
		If ([string]::IsNullOrEmpty($ShareForceSalting) -Or (-Not [string]($ShareForceSalting -as [int])))	{
			LogTrace ("ERROR: Parameter 'ShareForceSalting' (CSV file as input) is empty or incorrect. Please enter a value for this parameter in the CSV file")
			Write-Host "ERROR: Parameter 'ShareForceSalting' (CSV file as input) is empty or incorrect. Please enter a value for this parameter in the CSV file"
			Continue }

		### CONNECTION to vCenter
		LogTrace ("Connection to vCenter $vCenter")
		Write-Host "Connection to vCenter " -NoNewLine
		Write-Host "$vCenter... " -ForegroundColor Yellow -NoNewLine

		$rccnx = Connect-VIServer -Server $vCenter -Protocol https -Credential $Credential -WarningAction 0
		$topCnxVcc = "0"
		If ($rccnx -ne $null) {If ($rccnx.Isconnected) {$topCnxVcc = "1"	}	}

		If ($topCnxVcc -ne "1") {
			LogTrace ("[ERROR] Connection KO to vCenter $vCenter => Script halted")
			Write-Host "[NOK]`r`n" -ForegroundColor White -BackgroundColor Red
			$rc += 1
			Exit $rc	}
		Else {
			LogTrace ("[SUCCESS] Connection OK to vCenter $vCenter")
			Write-Host "[OK]`r`n" -ForegroundColor Black -BackgroundColor Green	}
			
		$MinTimeout = New-TimeSpan -Seconds $Maintenance_TimeOut_Delay | Select-Object -ExpandProperty Minutes
		$SecTimeout = New-TimeSpan -Seconds $Maintenance_TimeOut_Delay | Select-Object -ExpandProperty Seconds
		If ($SecTimeout -eq "0") {$SecTimeout = "00"}
			
		If ($ShareForceSalting -eq "0") { $Action = "ENABLING TPS"} Else { $Action = "DISABLING TPS"}
		If ($AllocGuestLargePage -eq "1") { $Action += ", LARGE PAGE FORCED"} Else { $Action += ", LARGE PAGE NOT FORCED"}
		
		$ESXNumOn = (Get-VMHost -Location $Cluster | Where {$_.PowerState -eq "PoweredOn"}).Count
		$ESXNumOff = (Get-VMHost -Location $Cluster | Where {$_.PowerState -eq "PoweredOff"}).Count
		If ($ESXNumOff -gt 0) { Write-Host "WARNING: $ESXNumOff ESXi of the cluster $Cluster is/are PoweredOff, it will not be performed" -ForegroundColor Red } Else { Write-Host "CAUTION: All ESXi of the cluster $Cluster are PoweredOn" -ForegroundColor Green }
		$ESXMaintenance = (Get-VMHost -Location $Cluster | Where {$_.ConnectionState -eq "Maintenance"}).Count
		If ($ESXMaintenance -gt 0) { Write-Host "WARNING: $ESXMaintenance ESXi of the cluster $Cluster is/are in maintenance mode, it will be performed`r`n" -ForegroundColor Red } Else { Write-Host "CAUTION: 0 ESXi of the cluster $Cluster is/are in maintenance mode`r`n" -ForegroundColor Green }
		
		### ESXi inventory in Cluster
		ForEach($ESX in (Get-VMHost -Location $Cluster | Where {$_.PowerState -eq "PoweredOn"} | Sort-Object Name)) {
			$StartTime = (Get-Date)
			[INT]$ElapsedTime = $NULL
		
			### Get current values
			LogTrace ("[INFO] $ESX [$Cluster] en cours de traitement")
			Write-Host "'$ESX' " -ForegroundColor Yellow -NoNewLine
			$ESXNumOn -= 1
			Write-Host "(#$ESXNumOn remain) " -ForegroundColor White -NoNewLine
			Write-Host "[$Cluster] " -NoNewLine
			Write-Host "(#$CountLine remain) " -ForegroundColor White -NoNewLine
			Write-Host "in progress..."
			LogTrace ("[INFO] APPLY action type to perform '$Action'")
			Write-Host "[INFO] APPLY action type to perform " -NoNewLine
			Write-Host "$Action" -ForegroundColor Red
			Write-Host "[INFO] GET current ESXi state " -NoNewLine
			$ESXMode = Get-VMHost -Name $ESX | Select-Object -ExpandProperty ConnectionState
			Write-Host "[$ESXMode]" -ForegroundColor DarkGray
			LogTrace ("[INFO] Current ESXi state '$ESXMode'")
			Write-Host "[INFO] GET current ESXi alarm action state " -NoNewLine
			$ESXAlarm = (Get-VMHost -Name $ESX).ExtensionData.AlarmActionsEnabled
			If ($ESXAlarm -eq "True") { $ESXAlarm = "Enabled" } Else { $ESXAlarm = "Disabled" }
			Write-Host "[$ESXAlarm]" -ForegroundColor DarkGray
			LogTrace ("[INFO] GET current ESXi alarm action state '$ESXAlarm'")
			Write-Host "[INFO] GET current parameter value " -NoNewLine
			Write-Host "'Mem.AllocGuestLargePage' " -NoNewLine -ForegroundColor Magenta
			$AllocGuestLargePage_Old = (Get-AdvancedSetting -Entity $ESX -Name "Mem.AllocGuestLargePage").Value
			Write-Host "[$AllocGuestLargePage_Old]" -ForegroundColor DarkGray
			LogTrace ("[INFO] GET current parameter value 'Mem.AllocGuestLargePage' is [$AllocGuestLargePage_Old]")
			Write-Host "[INFO] GET current parameter value " -NoNewLine
			Write-Host "'Mem.ShareForceSalting' " -NoNewLine -ForegroundColor Magenta
			$ShareForceSalting_Old = (Get-AdvancedSetting -Entity $ESX -Name "Mem.ShareForceSalting").Value
			Write-Host "[$ShareForceSalting_Old]" -ForegroundColor DarkGray
			LogTrace ("[INFO] GET current parameter value 'Mem.ShareForceSalting' actuel is [$ShareForceSalting_Old]")
			
			### SET new values
			### Test et/ou implémentation du paramètre 'Mem.AllocGuestLargePage'
			If ($AllocGuestLargePage_Old -ne $AllocGuestLargePage) {
			
				Write-Host "[INFO] " -NoNewLine
				Write-Host "--- Current LPAGE param NOT compliant with target [$AllocGuestLargePage], actions starting..." -ForegroundColor DarkYellow
				LogTrace ("[INFO] Current TPS parameters NOT compliant with target, actions starting...")
					
				### Arrêt des actions d'alarmes si actives
				If (((Get-VMHost -Name $ESX).ExtensionData.AlarmActionsEnabled) -eq "True") {
					Write-Host "[INFO] SET alarms actions mode to [Disabled]"
					LogTrace ("[INFO] SET alarms actions mode to [Disabled]")
					$alarmManager = Get-View AlarmManager -Server $vCenter
					$alarmManager.EnableAlarmActions((Get-VMHost -Name $ESX).Extensiondata.MoRef, $False) }
				Else {
					Write-Host "[INFO] SET alarms actions are already done [Disabled]"
					LogTrace ("[INFO] SET alarms actions are already done [Disabled]") }
				
				### Implémentation du paramètre
				Write-Host "[INFO] SET the parameter " -NoNewLine
				Write-Host "'Mem.AllocGuestLargePage' " -NoNewLine -ForegroundColor Magenta
				Write-Host "to " -NoNewLine
				Write-Host "[$AllocGuestLargePage] "  -NoNewLine -ForegroundColor DarkGray
				Get-AdvancedSetting -Entity $ESX -Name "Mem.AllocGuestLargePage" | Set-AdvancedSetting -Value $AllocGuestLargePage -Confirm:$False | Out-Null
				If((Get-AdvancedSetting -Entity $ESX -Name "Mem.AllocGuestLargePage").Value -ne $AllocGuestLargePage){ 
					$Resultat_LPAGE = "ERROR"
					LogTrace ("ERROR: 'Mem.AllocGuestLargePage' SET change was not applied to [$AllocGuestLargePage] on '$ESX'")
					Write-Host "[ERROR]" -ForegroundColor Red
					LogTrace ("[ERROR] SET the parameter 'Mem.AllocGuestLargePage' to [$AllocGuestLargePage]")	}
				Else {
					$Resultat_LPAGE = "SUCCESS"
					Write-Host "[SUCCESS]" -ForegroundColor Green
					LogTrace ("[SUCCESS] SET of the parameter 'Mem.AllocGuestLargePage' to [$AllocGuestLargePage]")	}
				
				### Mise en maintenance si ESXi connecté
				If ((Get-VMHost -Name $ESX | Select-Object -ExpandProperty ConnectionState) -eq "Connected") {
					Write-Host "[INFO] SET ESX into maintenance mode, it can take a long time, please wait..."
					LogTrace ("[INFO] SET ESX into maintenance mode")
					$Maintenance_Task = Set-VMHost -VMHost $ESX -State "Maintenance" -Server $vCenter -Confirm:$False -RunAsync:$True

					$Maintenance_StartTime = (Get-Date)
					$Maintenance_TimeOut = $False
					### Attente de la mise en maintenance
					Do {
						Start-Sleep -s 10
						$Maintenance_EndTime = (Get-Date)
						[INT]$ElapsedTime += 10
						Write-Host "[INFO] WAIT ESX maintenance mode for " -NoNewLine
						$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
						$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
						If ($Sec -eq "0") {$Sec = "00"}
						Write-Host "$($Min)min $($Sec)sec...`t" -ForegroundColor White -NoNewLine
						Write-Host "[Max. $($MinTimeout)min $($SecTimeout)sec]"
						LogTrace ("[INFO] WAIT ESX maintenance mode for $($Min)min $($Sec)sec...`t[Max. $($MinTimeout)min $($SecTimeout)sec]")
						If ($ElapsedTime -ge $Maintenance_TimeOut_Delay) {
							$Maintenance_TimeOut = $True
							Write-Host "[ERROR] KILL ESX maintenance mode task due to a timeout [$($MinTimeout)min $($SecTimeout)sec]" -ForegroundColor Red
							LogTrace ("[ERROR] KILL ESX maintenance mode task due to a timeout [$($MinTimeout)min $($SecTimeout)sec]")
							Stop-Task -Task $Maintenance_Task -Confirm:$False	}
					} Until (((Get-VMHost -Name $ESX | Select-Object -ExpandProperty ConnectionState) -eq "Maintenance") -OR ($Maintenance_TimeOut -eq $True))
					
					$Maintenance_EndTime = (Get-Date)
					$ElapsedTime = [math]::round((($Maintenance_EndTime-$Maintenance_StartTime).TotalSeconds), 0)
					$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
					$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
					If ($Sec -eq "0") {$Sec = "00"}
					
					If ((Get-VMHost -Name $ESX | Select-Object -ExpandProperty ConnectionState) -eq "Maintenance") {
						Write-Host "[INFO] ESX maintenance mode was done in " -NoNewLine
						Write-Host "$($Min)min $($Sec)sec...`t" -ForegroundColor White -NoNewLine
						Write-Host "[SUCCESS]" -ForegroundColor Green
						LogTrace ("[INFO] ESX maintenance mode was done in $($Min)min $($Sec)sec`t[SUCCESS]")	}
					Else {
						Write-Host "[ERROR] ESX maintenance mode was cancelled because timeout take over $($MinTimeout)min $($SecTimeout)sec" -ForegroundColor Red
						LogTrace ("[ERROR] ESX maintenance mode was cancelled because timeout take over $($MinTimeout)min $($SecTimeout)sec")	}
				}

				Else {
					Write-Host "[INFO] SET ESX into maintenance mode is already done"
					LogTrace ("[INFO] SET ESX into maintenance mode is already done") }
			}
			Else	{
				LogTrace ("[INFO] Parameter value 'Mem.AllocGuestLargePage' is already SET to [$AllocGuestLargePage]")
				Write-Host "[INFO] Parameter value " -NoNewLine
				Write-Host "'Mem.AllocGuestLargePage' " -NoNewLine -ForegroundColor Magenta
				Write-Host "is already SET to " -NoNewLine
				Write-Host "[$AllocGuestLargePage] " -ForegroundColor DarkGray -NoNewLine
				Write-Host "[OK]" -ForegroundColor Green }
			
			### Test et/ou implémentation du paramètre 'Mem.ShareForceSalting'
			If ($ShareForceSalting_Old -ne $ShareForceSalting)	{
				Write-Host "[INFO] " -NoNewLine
				Write-Host "--- Current SALTING param NOT compliant with target [$ShareForceSalting], actions starting..." -ForegroundColor DarkYellow
				LogTrace ("[INFO] Current SALTING parameters NOT compliant with target, actions starting...")
			
				### Arrêt des actions d'alarmes si actives
				If (((Get-VMHost -Name $ESX).ExtensionData.AlarmActionsEnabled) -eq "True") {
					LogTrace ("[INFO] SET alarms actions mode to [Disabled]")
					Write-Host "[INFO] SET alarms actions mode to [Disabled]"
					$alarmManager = Get-View AlarmManager -Server $vCenter
					$alarmManager.EnableAlarmActions((Get-VMHost -Name $ESX).Extensiondata.MoRef, $False) }
				Else {
					Write-Host "[INFO] SET alarms actions are already done [Disabled]"
					LogTrace ("[INFO] SET alarms actions are already done [Disabled]")	}
				
				### Implémentation du paramètre
				Write-Host "[INFO] SET the parameter " -NoNewLine
				Write-Host "'Mem.ShareForceSalting' " -NoNewLine -ForegroundColor Magenta
				Write-Host "to " -NoNewLine
				Write-Host "[$ShareForceSalting] " -NoNewLine -ForegroundColor DarkGray
				Get-AdvancedSetting -Entity $ESX -Name "Mem.ShareForceSalting" | SET-AdvancedSetting -Value $ShareForceSalting -Confirm:$False | Out-Null
				If((Get-AdvancedSetting -Entity $ESX -Name "Mem.ShareForceSalting").Value -ne "$ShareForceSalting"){
					$Resultat_SALTING = "ERROR"
					LogTrace ("[ERROR] SET the parameter 'Mem.ShareForceSalting' to [$ShareForceSalting]")
					Write-Host "[ERROR]" -ForegroundColor Red	}
				Else {
					$Resultat_SALTING = "SUCCESS"
					LogTrace ("[SUCCESS] SET the parameter 'Mem.ShareForceSalting' to [$ShareForceSalting]")
					Write-Host "[SUCCESS]" -ForegroundColor Green	}
				
				### Mise en maintenance si ESXi connecté
				If ((Get-VMHost -Name $ESX | Select-Object -ExpandProperty ConnectionState) -eq "Connected") {
					Write-Host "[INFO] SET into ESX maintenance mode, it can take a long time, please wait..."
					LogTrace ("[INFO] SET into ESX maintenance mode")
					$Maintenance_Task = Set-VMHost -VMHost $ESX -State "Maintenance" -Server $vCenter -Confirm:$False -RunAsync:$True
					
					$Maintenance_StartTime = (Get-Date)
					$Maintenance_TimeOut = $False
					### Attente de la mise en maintenance
					Do {
						Start-Sleep -s 10
						$Maintenance_EndTime = (Get-Date)
						[INT]$ElapsedTime += 10
						Write-Host "[INFO] WAIT ESX maintenance mode for " -NoNewLine
						$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
						$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
						If ($Sec -eq "0") {$Sec = "00"}
						Write-Host "$($Min)min $($Sec)sec...`t" -ForegroundColor White -NoNewLine
						Write-Host "[Max. $($MinTimeout)min $($SecTimeout)sec]"
						LogTrace ("[INFO] WAIT ESX maintenance mode for $($Min)min $($Sec)sec...`t[Max. $($MinTimeout)min $($SecTimeout)sec]")
						If ($ElapsedTime -ge $Maintenance_TimeOut_Delay) {
							$Maintenance_TimeOut = $True
							Write-Host "[ERROR] KILL ESX maintenance mode task due to a timeout [$($MinTimeout)min $($SecTimeout)sec]" -ForegroundColor Red
							LogTrace ("[ERROR] KILL ESX maintenance mode task due to a timeout [$($MinTimeout)min $($SecTimeout)sec]")
							Stop-Task -Task $Maintenance_Task -Confirm:$False	}
					} Until (((Get-VMHost -Name $ESX | Select-Object -ExpandProperty ConnectionState) -eq "Maintenance") -OR ($Maintenance_TimeOut -eq $True))
					$Maintenance_EndTime = (Get-Date)
					$ElapsedTime = [math]::round((($Maintenance_EndTime-$Maintenance_StartTime).TotalSeconds), 0)
					$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
					$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
					If ($Sec -eq "0") {$Sec = "00"}
					
					If ((Get-VMHost -Name $ESX | Select-Object -ExpandProperty ConnectionState) -eq "Maintenance") {
						Write-Host "[INFO] ESX maintenance mode was done in " -NoNewLine
						Write-Host "$($Min)min $($Sec)sec...`t" -ForegroundColor White -NoNewLine
						Write-Host "[SUCCESS]" -ForegroundColor Green
						LogTrace ("[INFO] ESX maintenance mode was done in $($Min)min $($Sec)sec`t[SUCCESS]")	}
					Else {
						Write-Host "[ERROR] ESX maintenance mode was cancelled because timeout take over $($MinTimeout)min $($SecTimeout)sec" -ForegroundColor Red
						LogTrace ("[ERROR] ESX maintenance mode was cancelled because timeout take over $($MinTimeout)min $($SecTimeout)sec")	}
				}
					
				Else {
					Write-Host "[INFO] SET ESX into maintenance mode is already done"
					LogTrace ("[INFO] SET ESX into maintenance mode is already done")	}
			}
			Else	{
				LogTrace ("[INFO] Parameter value 'Mem.ShareForceSalting' is already SET to [$ShareForceSalting]")
				Write-Host "[INFO] Parameter value " -NoNewLine
				Write-Host "'Mem.ShareForceSalting' " -NoNewLine -ForegroundColor Magenta
				Write-Host "is already SET to "-NoNewLine
				Write-Host "[$ShareForceSalting] " -ForegroundColor DarkGray -NoNewLine
				Write-Host "[OK]" -ForegroundColor Green }
			
			### Retour au mode de fonctionnement nominal
			If ($ESXMode -eq "Connected" -AND (Get-VMHost -Name $ESX | Select-Object -ExpandProperty ConnectionState) -ne "Connected") {
				Write-Host "[INFO] Quit ESX maintenance mode, please wait..."
				LogTrace ("[INFO] Quit ESX maintenance mode, please wait...")
				$Maintenance_Task = Set-VMHost -VMHost $ESX -State "Connected" -Server $vCenter -Confirm:$False | Out-Null }
			Else {
				Write-Host "[INFO] No action regarding ESX mode is required [-]"
				LogTrace ("[INFO] No action regarding ESX mode is required [-]")
				$Resultat = "SUCCESS" }
			
			If ($ESXAlarm -eq "Enabled" -AND ((Get-VMHost -Name $ESX).ExtensionData.AlarmActionsEnabled) -ne "True") {
				Write-Host "[INFO] SET alarms action mode to [Enabled]"
				LogTrace ("[INFO] SET alarms action mode to [Enabled]")
				$alarmManager = Get-View AlarmManager -Server $vCenter
				$alarmManager.EnableAlarmActions((Get-VMHost -Name $ESX).Extensiondata.MoRef, $True) }
			Else {
				Write-Host "[INFO] No action regarding alarms actions is required [-]"
				LogTrace ("[INFO] No action regarding alarms actions is required [-]")
				$Resultat = "SUCCESS" }
			
			### Affichage du résultat final
			Write-Host "[INFO] 'LPAGE' " -NoNewLine
			Write-Host "[$Resultat_LPAGE]" -NoNewLine -ForegroundColor Red
			Write-Host ", 'SALTING' " -NoNewLine
			Write-Host "[$Resultat_SALTING]" -ForegroundColor Red
			LogTrace ("[INFO] 'LPAGE' [$Resultat_LPAGE], 'SALTING' [$Resultat_SALTING]")
			
			### Affichage du temps de traitement
			$EndTime = (Get-Date)
			[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
			$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
			$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
			If ($Sec -eq "0") {$Sec = "00"}
			Write-Host "[TIME] Time elapsed: " -NoNewLine
			Write-Host "$($Min)min $($Sec)sec..." -ForegroundColor White
			LogTrace ("[TIME] Time elapsed: $($Min)min $($Sec)sec....")
			
			Write-Host "`r`n"
			LogTrace ("[FINISH] End processing" + $vbcrlf)
			
			### Write CSV outfile
			$LOG_Line = "$ESX;$vCenter;$Cluster;$Action;$Resultat_LPAGE;$Resultat_SALTING;$AllocGuestLargePage_Old;$ShareForceSalting_Old;$AllocGuestLargePage;$ShareForceSalting"
			Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
		}
	}
}

Else {
	### Exit Loop
	Write-Host "[INFO] Program cancelled by user !" -ForegroundColor White -BackgroundColor Red
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject "[INFO] Program cancelled by user !" -Append	}

### Exit program
Write-Host "Program finished !" -ForegroundColor Green
Out-File -Filepath $FicLog -Encoding UTF8 -InputObject "[INFO] Yippee! Program finished !" -Append