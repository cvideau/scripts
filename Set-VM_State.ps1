### V1.0 (octobre 2018) - Développeur: Christophe VIDEAU

Add-PSsnapin VMware.VimAutomation.Core
Clear-Host
Write-Host "Developped (2018) by Christophe VIDEAU - Version 1.0`r`n" -ForegroundColor White


$global:VM_FQDN						= $NULL
$global:VM_Error 					= 0
$global:VM_Success 					= 0
$global:StopStartVM_TimeOut_Delay 	= 120	### Timeout delay in seconds
$StartTime_Script 					= (Get-Date)
$LeftTimeEstimated 					= $NULL
$ProcessMode						= "Quick"	# Choose between 'Quick' or 'Normal'


### FUNCTION: SHUTDOWN VM
Function SHUTDOWN-VM($VM, $vCenter){
	Stop-VMGuest -VM $VMName -Confirm:$False -Server $vCenter | Out-Null
	Write-Host "[WARN] '$VM_FQDN' PowerOffGuest is processing !" -ForegroundColor Yellow
	If ($ProcessMode -eq "Normal")	{
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' PowerOffGuest is processing !"
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		
		$Stop_StartTime = (Get-Date)
		$Stop_TimeOut = $False
		[INT]$ElapsedTime = $NULL
		### Waiting stop VM
		Do {
			Sleep 5
			$Stop_EndTime = (Get-Date)
			[INT]$ElapsedTime += 5
			Write-Host "[INFO] '$VM_FQDN' wait PowerOffGuest for " -NoNewLine
			$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
			$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
			If ($Sec -eq "0") {$Sec = "00"}
			If ($Sec -eq "5") {$Sec = "05"}
			Write-Host "$($Min)min. $($Sec)sec...`t" -ForegroundColor White -NoNewLine
			Write-Host "[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' wait halt for $($Min)min. $($Sec)sec...`t[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			If ($ElapsedTime -ge $StopStartVM_TimeOut_Delay) {
				$Stop_TimeOut = $True
				Write-Host "[FAIL] '$VM_FQDN' PowerOffGuest task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]" -ForegroundColor Red
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' PowerOffGuest task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
		} Until (((Get-VM $VMName -Server $vCenter | Select PowerState).PowerState -eq "PoweredOff") -OR ($Stop_TimeOut -eq $True))
						
		$Stop_EndTime = (Get-Date)
		$ElapsedTime = [math]::round((($Stop_EndTime-$Stop_StartTime).TotalSeconds), 0)
		$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
		$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
		If ($Sec -eq "0") {$Sec = "00"}
		If ($Sec -eq "5") {$Sec = "05"}
						
		If ((Get-VM $VMName -Server $vCenter | Select PowerState).PowerState -eq "PoweredOff") {
			$VM_Success += 1
			Write-Host "[SUCC] '$VM_FQDN' PowerOffGuest was done in " -NoNewLine
			Write-Host "$($Min)min. $($Sec)sec..." -ForegroundColor White
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' PowerOffGuest was done in $($Min)min. $($Sec)sec."
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
		Else {
			$VM_Error += 1
			Write-Host "[FAIL] '$VM_FQDN' PowerOffGuest was take more than timeout delay. Over than $($MinTimeout)min. $($SecTimeout)sec." -ForegroundColor Red
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' PowerOffGuest was take more than timeout delay. Over than $($MinTimeout)min. $($SecTimeout)sec."
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
	}
}


### FUNCTION: POWER OFF VM
Function POWEROFF-VM($VM, $vCenter){
	Stop-VM -VM $VMName -Confirm:$False -RunAsync -Server $vCenter | Out-Null
	Write-Host "[WARN] '$VM_FQDN' PowerOff is processing !" -ForegroundColor Yellow
	If ($ProcessMode -eq "Normal")	{
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' PowerOff is processing !"
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		
		$Stop_StartTime = (Get-Date)
		$Stop_TimeOut = $False
		[INT]$ElapsedTime = $NULL
		### Waiting stop VM
		Do {
			Sleep 5
			$Stop_EndTime = (Get-Date)
			[INT]$ElapsedTime += 5
			Write-Host "[INFO] '$VM_FQDN' wait PowerOff for " -NoNewLine
			$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
			$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
			If ($Sec -eq "0") {$Sec = "00"}
			If ($Sec -eq "5") {$Sec = "05"}
			Write-Host "$($Min)min. $($Sec)sec...`t" -ForegroundColor White -NoNewLine
			Write-Host "[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' wait PowerOff for $($Min)min. $($Sec)sec...`t[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			If ($ElapsedTime -ge $StopStartVM_TimeOut_Delay) {
				$Stop_TimeOut = $True
				Write-Host "[FAIL] '$VM_FQDN' PowerOff task timeout, a forced PowerOff is processing [$($MinTimeout)min. $($SecTimeout)sec.]" -ForegroundColor Red
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' PowerOff task timeout, a forced PowerOff is processing [$($MinTimeout)min. $($SecTimeout)sec.]"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				Stop-VM -VM $VMName -Kill -Confirm:$False -Server $vCenter | Out-Null			}
		} Until (((Get-VM $VMName -Server $vCenter | Select PowerState).PowerState -eq "PoweredOff") -OR ($Stop_TimeOut -eq $True))
						
		$Stop_EndTime = (Get-Date)
		$ElapsedTime = [math]::round((($Stop_EndTime-$Stop_StartTime).TotalSeconds), 0)
		$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
		$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
		If ($Sec -eq "0") {$Sec = "00"}
		If ($Sec -eq "5") {$Sec = "05"}
						
		If ((Get-VM $VMName -Server $vCenter | Select PowerState).PowerState -eq "PoweredOff") {
			$VM_Success += 1
			Write-Host "[SUCC] '$VM_FQDN' PowerOff was done in " -NoNewLine
			Write-Host "$($Min)min. $($Sec)sec..." -ForegroundColor White
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' PowerOff was done in $($Min)min. $($Sec)sec."
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
		Else {
			$VM_Error += 1
			Write-Host "[FAIL] '$VM_FQDN' PowerOff was take more than timeout delay. Over than $($MinTimeout)min. $($SecTimeout)sec." -ForegroundColor Red
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' PowerOff was take more than timeout delay. Over than $($MinTimeout)min. $($SecTimeout)sec."
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
	}
}


### FUNCTION: START VM
Function ALIVE-VM($VM, $vCenter){
	Start-VM -VM $VMName -Confirm:$False -RunAsync -Server $vCenter | Out-Null
	Write-Host "[WARN] '$VM_FQDN' PowerOn is processing..." -ForegroundColor Yellow
	If ($ProcessMode -eq "Normal")	{
		Write-Host "[WARN] '$VM_FQDN' PowerOn is processing..." -ForegroundColor Yellow
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' PowerOn is processing..."
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		
		$Start_StartTime = (Get-Date)
		$Start_TimeOut = $False
		[INT]$ElapsedTime = $NULL
		### Waiting start VM
		Do {
			$ToolStatus = (Get-vm $VMName -Server $vCenter | Get-View).Guest.ToolsRunningStatus
			Sleep 5
			$Start_EndTime = (Get-Date)
			[INT]$ElapsedTime += 5
			Write-Host "[INFO] '$VM_FQDN' wait PowerOn for " -NoNewLine
			$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
			$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
			If ($Sec -eq "0") {$Sec = "00"}
			If ($Sec -eq "5") {$Sec = "05"}
			Write-Host "$($Min)min. $($Sec)sec...`t" -ForegroundColor White -NoNewLine
			Write-Host "[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' wait PowerOn for $($Min)min. $($Sec)sec...`t[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			If ($ElapsedTime -ge $StopStartVM_TimeOut_Delay) {
				$Start_TimeOut = $True
				Write-Host "[FAIL] '$VM_FQDN' PowerOn task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]" -ForegroundColor Red
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' PowerOn task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
		} Until (($ToolStatus -eq "guestToolsRunning") -OR ($Start_TimeOut -eq $True))
						
		$Start_EndTime = (Get-Date)
		$ElapsedTime = [math]::round((($Start_EndTime-$Start_StartTime).TotalSeconds), 0)
		$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
		$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
		If ($Sec -eq "0") {$Sec = "00"}
		If ($Sec -eq "5") {$Sec = "05"}
						
		If ((Get-VM $VMName -Server $vCenter | Select PowerState).PowerState -eq "PoweredOn") {
			$VM_Success += 1
			Write-Host "[SUCC] '$VM_FQDN' PowerOn was done in " -NoNewLine
			Write-Host "$($Min)min. $($Sec)sec...`t" -ForegroundColor White
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' PowerOn was done in $($Min)min. $($Sec)sec."
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
		Else {
			$VM_Error += 1
			Write-Host "[FAIL] '$VM_FQDN' was not have been processing due a PowerOn timeout. Over than $($MinTimeout)min. $($SecTimeout)sec." -ForegroundColor Red
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' was not have been processing due a PowerOn timeout. Over than $($MinTimeout)min. $($SecTimeout)sec."
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
	}
}


### FUNCTION: RESTART GUEST OS
Function RESTART-GUEST($VM, $vCenter){
	Restart-VMGuest -VM $VMName -Confirm:$False -Server $vCenter | Out-Null
	Write-Host "[WARN] '$VM_FQDN' RebootGuest is processing..." -ForegroundColor Yellow
	If ($ProcessMode -eq "Normal")	{
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' RebootGuest is processing..."
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		
		$Start_StartTime = (Get-Date)
		$Start_TimeOut = $False
		[INT]$ElapsedTime = $NULL
		### Waiting restart guest OS
		Do {
			$ToolStatus = (Get-vm $VMName -Server $vCenter | Get-View).Guest.ToolsRunningStatus
			Sleep 5
			$Start_EndTime = (Get-Date)
			[INT]$ElapsedTime += 5
			Write-Host "[INFO] '$VM_FQDN' wait RebootGuest for " -NoNewLine
			$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
			$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
			If ($Sec -eq "0") {$Sec = "00"}
			If ($Sec -eq "5") {$Sec = "05"}
			Write-Host "$($Min)min. $($Sec)sec...`t" -ForegroundColor White -NoNewLine
			Write-Host "[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' RebootGuest restart for $($Min)min. $($Sec)sec...`t[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			If ($ElapsedTime -ge $StopStartVM_TimeOut_Delay) {
				$Start_TimeOut = $True
				Write-Host "[FAIL] '$VM_FQDN' RebootGuest task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]" -ForegroundColor Red
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' RebootGuest task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
		} Until (($ToolStatus -eq "guestToolsRunning") -OR ($Start_TimeOut -eq $True))
						
		$Start_EndTime = (Get-Date)
		$ElapsedTime = [math]::round((($Start_EndTime-$Start_StartTime).TotalSeconds), 0)
		$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
		$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
		If ($Sec -eq "0") {$Sec = "00"}
		If ($Sec -eq "5") {$Sec = "05"}
						
		If ((Get-VM $VMName -Server $vCenter | Select PowerState).PowerState -eq "PoweredOn") {
			$VM_Success += 1
			Write-Host "[SUCC] '$VM_FQDN' RebootGuest was done in " -NoNewLine
			Write-Host "$($Min)min. $($Sec)sec...`t" -ForegroundColor White
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' RebootGuest was done in $($Min)min. $($Sec)sec."
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
		Else {
			$VM_Error += 1
			Write-Host "[FAIL] '$VM_FQDN' was not have been processing due a RebootGuest timeout. Over than $($MinTimeout)min. $($SecTimeout)sec." -ForegroundColor Red
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' was not have been processing due a RebootGuest timeout. Over than $($MinTimeout)min. $($SecTimeout)sec."
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
	}
}


### FUNCTION: REINIT VM
Function REINIT-VM($VM, $vCenter){
	(Get-VM -Name $VMName).ExtensionData.ResetVM()
	Write-Host "[WARN] '$VM_FQDN' Reinit is processing..." -ForegroundColor Yellow
	If ($ProcessMode -eq "Normal")	{
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' Reinit VM is processing..."
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		
		$Start_StartTime = (Get-Date)
		$Start_TimeOut = $False
		[INT]$ElapsedTime = $NULL
		### Waiting reinit VM
		Do {
			$ToolStatus = (Get-vm $VMName -Server $vCenter | Get-View).Guest.ToolsRunningStatus
			Sleep 5
			$Start_EndTime = (Get-Date)
			[INT]$ElapsedTime += 5
			Write-Host "[INFO] '$VM_FQDN' wait boot for " -NoNewLine
			$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
			$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
			If ($Sec -eq "0") {$Sec = "00"}
			If ($Sec -eq "5") {$Sec = "05"}
			Write-Host "$($Min)min. $($Sec)sec...`t" -ForegroundColor White -NoNewLine
			Write-Host "[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' wait Reinit for $($Min)min. $($Sec)sec...`t[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			If ($ElapsedTime -ge $StopStartVM_TimeOut_Delay) {
				$Start_TimeOut = $True
				Write-Host "[FAIL] '$VM_FQDN' Reinit task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]" -ForegroundColor Red
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Reinit task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
		} Until (($ToolStatus -eq "guestToolsRunning") -OR ($Start_TimeOut -eq $True))
						
		$Start_EndTime = (Get-Date)
		$ElapsedTime = [math]::round((($Start_EndTime-$Start_StartTime).TotalSeconds), 0)
		$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
		$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
		If ($Sec -eq "0") {$Sec = "00"}
		If ($Sec -eq "5") {$Sec = "05"}
						
		If ((Get-VM $VMName -Server $vCenter | Select PowerState).PowerState -eq "PoweredOn") {
			$VM_Success += 1
			Write-Host "[SUCC] '$VM_FQDN' Reinit was done in " -NoNewLine
			Write-Host "$($Min)min. $($Sec)sec...`t" -ForegroundColor White
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' Reinit was done in $($Min)min. $($Sec)sec."
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
		Else {
			$VM_Error += 1
			Write-Host "[FAIL] '$VM_FQDN' was not have been processing due a Reinit timeout. Over than $($MinTimeout)min. $($SecTimeout)sec." -ForegroundColor Red
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' was not have been processing due a Reinit timeout. Over than $($MinTimeout)min. $($SecTimeout)sec."
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
	}
}


### FUNCTION: Create and append LOG file
Function LogTrace ($Message){
	$Message = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " " + $Message
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Message -Append }

### LOG file initialization
$dat	= Get-Date -UFormat "%Y_%m_%d_%H_%M_%S"
$vbcrlf		= "`r`n"
$ScriptName	= [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition)

$RepLog     = $PathScript + "LOG"
If (!(Test-Path $RepLog)) { New-Item -Path $PathScript -ItemType Directory -Name "LOG"}
$RepLog    					= $RepLog + "\"				### Script folder
$Random						= Get-Random -Max 1000		### ID number ramdom
$FicLog     				= $RepLog + $ScriptName + "-" + $dat + "_(" + $env:UserName + ")_ID" + $Random + ".log"		### Script LOG folder
$FicRes     				= $RepLog + $ScriptName + "-" + $dat + "_(" + $env:UserName + ")_ID" + $Random + ".csv"		### Script CSV folder
$LineSep    				= "=" * 70
$MinTimeout = New-TimeSpan -Seconds $StopStartVM_TimeOut_Delay | Select-Object -ExpandProperty Minutes		### $StopStartVM_TimeOut_Delay conversion in minutes
$SecTimeout = New-TimeSpan -Seconds $StopStartVM_TimeOut_Delay | Select-Object -ExpandProperty Seconds		### $StopStartVM_TimeOut_Delay conversion in seconds
If ($SecTimeout -eq "0") {$SecTimeout = "00"}

$Line = "## BEGIN Script change VM state ##"
If (!(Test-Path $FicLog)) {
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " " + $Line
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Line -Append
} Else {
	LogTrace ($Line)
}
Write-Host "LogFile initialization is done !" -ForegroundColor Green
LogTrace ($LineSep)

### Load assembly & context
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
#$ErrorActionPreference= 'SilentlyContinue'
Write-Host "Setting Execution Policy to Unrestricted" -ForegroundColor Green
Set-ExecutionPolicy Unrestricted -Force | Out-Null
### https://kb.vmware.com/selfservice/microsites/search.do?language=en_US&cmd=displayKC&externalId=2009857
Write-Host "[INFO]:Information [WARN]:Warning [FAIL]:Failed [SUCC]:Success`r`n" -ForegroundColor White

### TEST CSV file exist
<# If (-NOT (Test-Path .\VMList.csv)) {
	### EXIT Program
	[System.Windows.Forms.MessageBox]::Show("CSV file (.\VMList.csv) is missing, program is aborted !", "Abort" ,0 , 16) | Out-Null
	Write-Host "Program is aborted !" -ForegroundColor White -BackgroundColor Red
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] CSV file is missing, program aborted !"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	Continue } #>

If (-NOT ($args[0])) {
	[System.Windows.Forms.MessageBox]::Show("CSV input file parameter is missing... Please specify a CSV input file as parameter. Program is aborted !", "Abort" ,0 , 16) | Out-Null
	Write-Host "[FAIL] CSV input file parameter is missing... " -ForegroundColor White -BackgroundColor Red
	Write-Host "Please specify a CSV input file as parameter. Program is aborted !" -ForegroundColor White -BackgroundColor Red
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] CSV input file parameter is missing... Please specify a CSV input file as parameter. Program is aborted !"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	Write-Host -NoNewLine "Press any key to quit the program...`r`n"
	$NULL = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	Exit
}

### DEFINE outfile column header
$LOG_Head  = "Date;VM;OS;vCenter;Cluster;vmTools;[T0]State;[T1]State;Details"
Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Head -Append

$Reponse = [System.Windows.Forms.MessageBox]::Show("Are you sure that you want to change the VM state on each VM presents in the input file ?", "Confirmation" , 4, 32)
If ($Reponse -eq "Yes") {

	$VMList = Import-CSV -Delimiter ';' .\$($args[0])
	$CountLine = Get-ChildItem -Path .\$($args[0]) -Recurse | Get-Content | Measure-Object -Line | Select -Expand Lines
	ForEach ($Item in $VMList) {
		$LeftTimeEstimated_Start 	= Get-Date												### DateTime each VM VM start
		$VMName						= $Item.("Hostname")									### VM
		$vCenter					= $Item.("vCenter")										### vCenter
		$Action						= $Item.("Action")										### Action type
		#[System.Windows.Forms.MessageBox]::Show("$VMName $RAM_Target")

		### USER identification initialization
		$PathScript = ($myinvocation.MyCommand.Definition | Split-Path -parent) + "\"
		$user       = "CSP_SCRIPT_ADM"
		$fickey     = "D:\Scripts\Credentials\key.crd "
		$ficcred    = "D:\Scripts\Credentials\vmware_adm.crd"
		$key        = get-content $fickey
		$pwd        = Get-Content $ficcred | ConvertTo-SecureString -key $key
		$Credential = New-Object System.Management.Automation.PSCredential $user, $pwd
		
		### CSV file CHECK
		If ([String]::IsNullOrEmpty($VMName))	{	LogTrace ("[FAIL] Parameter 'Hostname' is empty. Please enter the name of the target 'VM' in the CSV file." + $vbcrlf);	Continue }
		If ([String]::IsNullOrEmpty($vCenter))	{	LogTrace ("[FAIL] Parameter 'vCenter' is empty. Please enter the name of the target 'vCenter' in the CSV file." + $vbcrlf);	Continue }
		If ([String]::IsNullOrEmpty($Action))	{	LogTrace ("[FAIL] Parameter 'Action' is empty or incorrect. Please enter an action type for this parameter in the CSV file." + $vbcrlf);	Continue }
				
		### CONNECTION to vCenter
		If ($vCenter -ne $vCenter_Previous)	{
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] Connection to vCenter '$vCenter'..."
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			Write-Host "Connection to vCenter " -NoNewLine
			Write-Host "'$vCenter'... " -ForegroundColor Yellow -NoNewLine
			Write-Host "is processing" -ForegroundColor Green
		
			$rccnx = Connect-VIServer -Server $vCenter -Protocol https -Credential $Credential -WarningAction 0
			$topCnxVcc = "0"
			If ($rccnx -ne $null) {
				If ($rccnx.Isconnected) {
					$topCnxVcc = "1"	}}

			If ($topCnxVcc -ne "1") {
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] Connection KO to vCenter '$vCenter' => Script halted"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				Write-Host "[FAIL] Connection KO to vCenter '$vCenter' => Script halted" -ForegroundColor White -BackgroundColor Red
				$rc += 1
				[INT]$VM_Error += 1
				Continue	}
			Else {
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] Connection OK to vCenter '$vCenter'"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				Write-Host "[SUCC] Connection OK to vCenter '$vCenter'" -ForegroundColor Black -BackgroundColor Green	}
		}
		
		$vCenter_Previous = $vCenter

		If ($ProcessMode -eq "Normal")	{
			$StartTime = (Get-Date)
			$StartTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
			$Cluster	= Get-Cluster -VM $VMName -Server $vCenter | Select -Expand Name
			$VM_OS		= (Get-View (Get-VM $VMName)).Guest.GuestFullName	### VM OS
		}
		$VM_FQDN	= (Get-VM $VMName).Guest.HostName						### VM FQDN
		
		### VMTools CHECK
		$State = Get-View -ViewType VirtualMachine -Filter @{"Name" = $VMName}
		If ($State.Guest.ToolsRunningStatus -ne "guestToolsRunning") { $VM_FQDN = $VMName; $VM_OS = "Unknown" }

		If ($ProcessMode -eq "Normal")	{
			If ($LeftTimeEstimated -ne $NULL) {
				$Min = New-TimeSpan -Seconds $($LeftTimeEstimated * $($CountLine - 1)) | Select-Object -ExpandProperty Minutes
				$Sec = New-TimeSpan -Seconds $($LeftTimeEstimated * $($CountLine - 1)) | Select-Object -ExpandProperty Seconds
				If ($Sec -eq "0") {$Sec = "00"}
				Write-Host "[INFO] $($CountLine - 1) VM remaining to process, the remaining time is estimated at $($Min)min. $($Sec)sec. Next...`r`n" -ForegroundColor Yellow
			} Else { Write-Host "[INFO] $($CountLine - 1) VM remaining to process, the estimated remaining time is still unknown. Next...`r`n" -ForegroundColor Yellow }
			$CountLine -= 1
			
			Write-Host "[INFO] '$VM_FQDN' guest OS name '$VM_OS'" -ForegroundColor Gray
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' guest OS name '$VM_OS'"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			Write-Host "[INFO] '$VM_FQDN' action start time on $StartTime_Display" -ForegroundColor White
			If ($State.Runtime.PowerState -eq "PoweredOn") {	Write-Host "[INFO] '$VM_FQDN' is 'Power On'" }	Else {	Write-Host "[INFO] '$VM_FQDN' is 'Power Off'"	}
			Write-Host "[INFO] '$VM_FQDN' is present on cluster '$Cluster'" -ForegroundColor Gray
		
			If ($State.Guest.ToolsRunningStatus -eq "guestToolsRunning") {
				$vmTools_Status = "Running"
				Write-Host "[INFO] '$VM_FQDN' vmTools are running, good news !" -ForegroundColor Green
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' VMTools are running, good news !"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			} Else {
				$vmTools_Status = "Not Running"
				$VM_FQDN = $VMName
				Write-Host "[WARN] '$VM_FQDN' <FQDN unknown> vmTools are NOT running, very bad news !" -ForegroundColor Red
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' <FQDN unknown> VMTools are NOT running, very bad news !"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
			
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' is present on cluster '$Cluster'"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		}		
		
		Switch ($Action)	{
			"PowerOn" 		{	If ($State.Runtime.PowerState -eq "PoweredOff")	{ ALIVE-VM -VM $VMName -vCenter $vCenter }
								Else {
									Write-Host "[WARN] '$VM_FQDN' VM is already powered on !" -ForegroundColor Red
									If ($ProcessMode -eq "Normal")	{
										$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [WARN] '$VM_FQDN' VM is already powered on !"
										Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
										$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;PoweredOff;PoweredOn;Done"
									}
									Continue	}
							}
			
			"PowerOff"		{ If ($State.Runtime.PowerState -eq "PoweredOn")	{ POWEROFF-VM -VM $VMName -vCenter $vCenter }
								Else {
									Write-Host "[WARN] '$VM_FQDN' VM is already powered off !" -ForegroundColor Red
									If ($ProcessMode -eq "Normal")	{
										$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [WARN] '$VM_FQDN' VM is already powered off !"
										Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
										$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;PoweredOn;PoweredOff;Done"
									}
									Continue	}
							}
			
			"Reinit"		{ If ($State.Runtime.PowerState -eq "PoweredOn")	{ REINIT-VM -VM $VMName -vCenter $vCenter }
								Else {
									Write-Host "[WARN] '$VM_FQDN' VM is already powered off !" -ForegroundColor Red
									If ($ProcessMode -eq "Normal")	{
										$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [WARN] '$VM_FQDN' VM is already powered off !"
										Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
										$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;PoweredOn;Reinit;Done"
									}
									Continue	}
							}
			
			"PowerOffGuest"	{ If (($State.Guest.ToolsRunningStatus -eq "guestToolsRunning") -and ($State.Runtime.PowerState -eq "PoweredOn")) { SHUTDOWN-VM -VM $VMName -vCenter $vCenter; Continue }
								Else {
									Write-Host "[FAIL] '$VM_FQDN' <FQDN unknown> vmTools (mandatory) are not detected or VM is powered off !" -ForegroundColor Red
									If ($ProcessMode -eq "Normal")	{
										$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' <FQDN unknown> vmTools (mandatory) are not detected or VM is powered off !"
										Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
										$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;PoweredOn;PoweredOff Guest;Done"
									}
									Continue	}
							}
			
			"RebootGuest"	{ If (($State.Guest.ToolsRunningStatus -eq "guestToolsRunning") -and ($State.Runtime.PowerState -eq "PoweredOn")) { RESTART-GUEST -VM $VMName -vCenter $vCenter; Continue }
								Else {
									Write-Host "[FAIL] '$VM_FQDN' <FQDN unknown> vmTools (mandatory) are not detected or VM is powered off !" -ForegroundColor Red
									If ($ProcessMode -eq "Normal")	{
										$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' <FQDN unknown> vmTools (mandatory) are not detected or VM is powered off !"
										Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
										$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;PoweredOn;Reboot guest;Done"
									}
									Continue	}
							}
		
			DEFAULT { exit}
		}
		
		If ($ProcessMode -eq "Normal")	{
			$EndTime = (Get-Date)
			$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
			$ElapsedTime = $NULL
			[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
			
			$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
			$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
			If ($Sec -eq "0") {$Sec = "00"}
			If ($Sec -eq "5") {$Sec = "05"}

			Write-Host "[INFO] '$VM_FQDN' action end time on $EndTime_Display" -ForegroundColor White
			Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
			
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				
			### Write CSV outfile
			$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Target;$($RAM_Value-$RAM_Target);$CPU_Value;$CPU_Target;$($CPU_Target-$CPU_Value);Operation is done"
			Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
			
			$LeftTimeEstimated_Finish = Get-Date
			$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
		}
	}
}
Else {
	### Exit Loop
	Write-Host "[INFO] Program cancelled by user !" -ForegroundColor White -BackgroundColor Red
	If ($ProcessMode -eq "Normal")	{
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] Program cancelled by user !"
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	}
}

$EndTime_Script = (Get-Date)
[String]$ElapsedTime_Script = [math]::round((($EndTime_Script-$StartTime_Script).TotalSeconds), 0)
$Min = New-TimeSpan -Seconds $ElapsedTime_Script | Select-Object -ExpandProperty Minutes
$Sec = New-TimeSpan -Seconds $ElapsedTime_Script | Select-Object -ExpandProperty Seconds
If ($Sec -eq "0") {$Sec = "00"}
If ($Sec -eq "5") {$Sec = "05"}

### Exit program
Write-Host "Program finished !" -ForegroundColor Red

If ($ProcessMode -eq "Normal")	{
	Write-Host "[INFO] $VM_Success VM succedeed, $VM_Error VM failed... (Completed in $($Min)min. $($Sec)sec.)" -ForegroundColor White
	$Line = "`r`n" + (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] $VM_Success VM succedeed, $VM_Error VM failed... (Completed in $($Min)min. $($Sec)sec.)`r`n" + (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] Yippee! Program is finished !`r"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
}
Write-Host -NoNewLine "Press any key to quit the program...`r`n"
$NULL = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")