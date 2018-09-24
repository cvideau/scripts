### V1.0 (juin 2018) - Développeur: Christophe VIDEAU

Add-PSsnapin VMware.VimAutomation.Core
Clear-Host
Write-Host "[BT122901] Copyright (2018) Christophe VIDEAU - Version 1.0`r`n" -ForegroundColor White

$global:State_MODIFY_RAM = $NULL
$global:State_MODIFY_CPU = $NULL

Function Start-Sleep($Seconds)	{
    $doneDT = (Get-Date).AddSeconds($seconds)
    while($doneDT -gt (Get-Date)) {
        $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
        $percent = ($seconds - $secondsLeft) / $seconds * 100
        Write-Progress -Activity "Sleeping" -Status "Sleeping..." -SecondsRemaining $secondsLeft -PercentComplete $percent
        [System.Threading.Thread]::Sleep(500)
    }
    Write-Progress -Activity "Sleeping" -Status "Sleeping..." -SecondsRemaining 0 -Completed
}


### FUNCTION: Enable the VM Hot Add RAM & CPU features
Function Enable-MemHotAdd($VM, $vCenter){
   	$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select MemoryHotAddEnabled
	If ($HotPlugStatus.MemoryHotAddEnabled -eq $False) {
		### ENABLE HOT add RAM
		$vmview = Get-vm $VM -Server $vCenter | Get-View 
		$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
		$extra = New-Object VMware.Vim.optionvalue
		$extra.Key="mem.hotadd"
		$extra.Value="true"
		$vmConfigSpec.extraconfig += $extra
		$vmview.ReconfigVM($vmConfigSpec)
		
		Write-Host "[INF] '$VM' hot add RAM activation parameter is initiating" -ForegroundColor White
		$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select MemoryHotAddEnabled
		If ($HotPlugStatus.MemoryHotAddEnabled -eq $False) {
			Write-Host "[ERR] '$VM' Error during process. Hot add RAM parameter activation has failed" -ForegroundColor Red
			$Line = (Get-Date -format G) + " " + "[ERR] '$VM' Error during process. Hot add RAM parameter activation has failed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		} Else {
			Write-Host "[SUC] '$VM' hot add RAM parameter activation has succedeed" -ForegroundColor Green
			$Line = (Get-Date -format G) + " " + "[SUC] '$VM' hot add RAM parameter activation has succedeed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		}
	}
	Else {
		Write-Host "[INF] '$VM' hot add RAM activation parameter is already activated, nothing to do" -ForegroundColor White
		$Line = (Get-Date -format G) + " " + "[INF] '$VM' hot add RAM parameter is already activated, nothing to do"
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	}
}

Function Enable-vCpuHotAdd($VM, $vCenter){
   	$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select CpuHotAddEnabled
	If ($HotPlugStatus.CpuHotAddEnabled -eq $False) {
		### ENABLE HOT add CPU
		$vmview = Get-vm $VM -Server $vCenter | Get-View 
		$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
		$extra = New-Object VMware.Vim.optionvalue
		$extra.Key="vcpu.hotadd"
		$extra.Value="true"
		$vmConfigSpec.extraconfig += $extra
		$vmview.ReconfigVM($vmConfigSpec)
		
		Write-Host "[INF] '$VM' hot add CPU activation parameter is initiating" -ForegroundColor White
		$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select CpuHotAddEnabled
		If ($HotPlugStatus.CpuHotAddEnabled -eq $False) {
			Write-Host "[ERR] '$VM' Error during process. Hot add CPU parameter activation has failed" -ForegroundColor Red
			$Line = (Get-Date -format G) + " " + "[ERR] '$VM' Error during process. Hot add CPU parameter activation has failed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		} Else {
			Write-Host "[SUC] '$VM' hot add CPU parameter activation has succedeed" -ForegroundColor Green
			$Line = (Get-Date -format G) + " " + "[SUC] '$VM' hot add CPU parameter activation has succedeed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		}
	}
	Else {
		Write-Host "[INF] '$VM' hot add CPU activation parameter is already activated, nothing to do" -ForegroundColor White
		$Line = (Get-Date -format G) + " " + "[INF] '$VM' hot add CPU parameter is already activated, nothing to do"
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	}
}


### FUNCTION: Disable the VM Hot Add RAM & CPU features
Function Disable-MemHotAdd($VM, $vCenter){
	$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select MemoryHotAddEnabled
	If ($HotPlugStatus.MemoryHotAddEnabled -eq $True) {
		### ENABLE HOT add RAM
		$vmview = Get-VM $VM -Server $vCenter | Get-View 
		$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
		$extra = New-Object VMware.Vim.optionvalue
		$extra.Key="mem.hotadd"
		$extra.Value="false"
		$vmConfigSpec.extraconfig += $extra
		$vmview.ReconfigVM($vmConfigSpec)
		
		Write-Host "[INF] '$VM' hot add RAM deactivation parameter is initiating" -ForegroundColor White
		$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select MemoryHotAddEnabled
		If ($HotPlugStatus.MemoryHotAddEnabled -eq $True) {
			Write-Host "[ERR] '$VM' Error during process. Hot add RAM parameter deactivation has failed" -ForegroundColor Red
			$Line = (Get-Date -format G) + " " + "[ERR] '$VM' Error during process. Hot add RAM parameter deactivation has failed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		} Else {
			Write-Host "[SUC] '$VM' hot add RAM parameter deactivation has succedeed" -ForegroundColor Green
			$Line = (Get-Date -format G) + " " + "[SUC] '$VM' hot add RAM parameter deactivation has succedeed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		}
	}
	Else {
		Write-Host "[INF] '$VM' hot add RAM deactivation parameter is already activated, nothing to do" -ForegroundColor White
		$Line = (Get-Date -format G) + " " + "[INF] '$VM' hot add RAM parameter is already activated, nothing to do"
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	}
}

Function Disable-vCpuHotAdd($VM, $vCenter){
   	$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select CpuHotAddEnabled
	If ($HotPlugStatus.CpuHotAddEnabled -eq $True) {
		$vmview = Get-vm $VM -Server $vCenter | Get-View 
		$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
		$extra = New-Object VMware.Vim.optionvalue
		$extra.Key="vcpu.hotadd"
		$extra.Value="false"
		$vmConfigSpec.extraconfig += $extra
		$vmview.ReconfigVM($vmConfigSpec)
		
		Write-Host "[INF] '$VM' hot add CPU activation parameter is initiating" -ForegroundColor White
		$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select CpuHotAddEnabled
		If ($HotPlugStatus.CpuHotAddEnabled -eq $True) {
			Write-Host "[ERR] '$VM' Error during process. Hot add CPU parameter activation has failed" -ForegroundColor Red
			$Line = (Get-Date -format G) + " " + "[ERR] '$VM' Error during process. Hot add CPU parameter activation has failed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		} Else {
			Write-Host "[SUC] '$VM' hot add CPU parameter activation has succedeed" -ForegroundColor Green
			$Line = (Get-Date -format G) + " " + "[SUC] '$VM' hot add CPU parameter activation has succedeed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		}
	}
	Else {
		Write-Host "[INF] '$VM' hot add CPU activation parameter is already activated, nothing to do" -ForegroundColor White
		$Line = (Get-Date -format G) + " " + "[INF] '$VM' hot add CPU parameter is already activated, nothing to do"
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	}
}
   

### FUNCTION: ADD RAM (at hot if possible)
Function MODIFY-RAM($VM, $vCenter){
Set-VM $VM -MemoryGB $RAM_Target -Confirm:$False | Out-Null
$State = Get-View -ViewType VirtualMachine -Filter @{"Name" = $VM}
If (($State.Config.Hardware.MemoryMB / 1024) -eq $RAM_Target) {
	Write-Host "[SUC] '$VM' has now '$RAM_Target'GB RAM" -ForegroundColor Green
	$Line = (Get-Date -format G) + " " + "[SUC] '$VM' has now '$RAM_Target'GB RAM"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	$State_MODIFY_RAM = "Success"	}
Else {
	Write-Host "[ERR] '$VM' Error during process. Doesn't perform RAM operation !" -ForegroundColor Red
	$Line = (Get-Date -format G) + " " + "[ERR] '$VM' Error during process. Doesn't perform RAM operation !"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	$State_MODIFY_RAM = "Error"	}
}


### FUNCTION: ADD VM NOTES AND VM TAG FOR RAM
Function MODIFY-RAM_INFOS($VM, $vCenter){
	### Add Notes
	$VMNotes = Get-VM $VM -Server $vCenter | Select-Object -ExpandProperty Notes
	Set-VM $VM -Notes "$VMNotes`r[BT122901-OPTIMISATION RAM] La capacite RAM a ete modifiee le $(Get-Date) de '$RAM_Value'Go a '$RAM_Target'Go" -Confirm:$False | Out-Null
	Write-Host "[INF] '$VM' a note has been added" -ForegroundColor Green
	$Line = (Get-Date -format G) + " " + "[INF] '$VM' a note has been added"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	
	If ($RAM_Value -lt $RAM_Target) {	$RAM_Tag = "VM_OPTIMIZE_ADD_RAM";	$RAM_Tag_Description = "[BT122901] Marqueur d'ajout de capacite RAM"	} Else {	$RAM_Tag = "VM_OPTIMIZE_SHRINK_RAM";	$RAM_Tag_Description = "[BT122901] Marqueur de réduction de capacite RAM"	}
	### CHECK if TAG exists
	$RAM_TAG_Exist = Get-Tag -Category "ADMINISTRATION" -Name $RAM_Tag -Server $vCenter -ErrorAction SilentlyContinue
	If (-NOT $RAM_TAG_Exist) {	New-Tag -Name $RAM_Tag -Category "ADMINISTRATION" -Description $RAM_Tag_Description | Out-Null	}
	
	### Add TAG
	$myTag = Get-Tag -Category "ADMINISTRATION" -Name $RAM_Tag -Server $vCenter | Out-Null
	Get-VM -Name $VM | New-TagAssignment -Tag $RAM_Tag | Out-Null
	Write-Host "[INF] '$VM' a TAG '$RAM_Tag' has been added" -ForegroundColor Green
	$Line = (Get-Date -format G) + " " + "[INF] '$VM' a TAG '$RAM_Tag' has been added"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
}


### FUNCTION: ADD CPU (at hot if possible)
Function MODIFY-CPU($VM, $vCenter){
Set-VM $VM -NumCpu $CPU_Target -Confirm:$False | Out-Null
$State = Get-View -ViewType VirtualMachine -Filter @{"Name" = $VM}
If ($State.Config.Hardware.NumCPU -eq $CPU_Target) {
	Write-Host "[SUC] '$VM' has now '$CPU_Target'CPU" -ForegroundColor Green
	$Line = (Get-Date -format G) + " " + "[SUC] '$VM' has now '$CPU_Target'CPU"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	$State_MODIFY_CPU = "Success"	}
Else {
	Write-Host "[ERR] '$VM' Error during process. Doesn't perform CPU operation !" -ForegroundColor Red
	$Line = (Get-Date -format G) + " " + "[ERR] '$VM' Error during process. Doesn't perform CPU operation !"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	$State_MODIFY_CPU = "Error"	}
}


### FUNCTION: ADD VM NOTES AND VM TAG FOR CPU
Function MODIFY-CPU_INFOS($VM, $vCenter){
	### Add Notes
	$VMNotes = Get-VM $VM -Server $vCenter | Select-Object -ExpandProperty Notes
	Set-VM $VM -Notes "$VMNotes`r[BT122901-OPTIMISATION CPU] Le nombre de CPU a ete modifie le $(Get-Date) de '$CPU_Value'CPU a '$CPU_Target'CPU" -Confirm:$False | Out-Null
	Write-Host "[INF] '$VM' a note has been added" -ForegroundColor Green
	$Line = (Get-Date -format G) + " " + "[INF] '$VM' a note has been added"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	
	If ($CPU_Value -lt $CPU_Target) {	$CPU_Tag = "VM_OPTIMIZE_ADD_CPU";	$CPU_Tag_Description = "[BT122901] Marqueur d'ajout de capacite CPU"	} Else {	$CPU_Tag = "VM_OPTIMIZE_SHRINK_CPU";	$CPU_Tag_Description = "[BT122901] Marqueur de réduction de capacite CPU"	}
	### CHECK if TAG exists
	$RAM_TAG_Exist = Get-Tag -Category "ADMINISTRATION" -Name $CPU_Tag -Server $vCenter -ErrorAction SilentlyContinue
	If (-NOT $RAM_TAG_Exist) {	New-Tag -Name $CPU_Tag -Category "ADMINISTRATION" -Description $CPU_Tag_Description | Out-Null	}
	
	### Add TAG
	$myTag = Get-Tag -Category "ADMINISTRATION" -Name $CPU_Tag -Server $vCenter | Out-Null
	Get-VM -Name $VM | New-TagAssignment -Tag $CPU_Tag | Out-Null
	Write-Host "[INF] '$VM' a TAG '$CPU_Tag' has been added" -ForegroundColor Green
	$Line = (Get-Date -format G) + " " + "[INF] '$VM' a TAG '$CPU_Tag' has been added"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
}


### FUNCTION: SHUTDOWN VM
Function SHUTDOWN-VM($VM, $vCenter){
	Write-Host "[WAR] '$VMName' will shutdown in 5 seconds... Press CTRL+C to CANCEL !" -ForegroundColor Yellow
	Start-Sleep 5
	Stop-VMGuest -VM $VMName -Confirm:$False -Server $vCenter | Out-Null
	Write-Host "[WAR] '$VMName' Warning: Shutdown is processing !" -ForegroundColor Yellow
	$Line = (Get-Date -format G) + " " + "[INF] '$VMName' shutdown is processing..."
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				
	Do {} until ((Get-VM $VMName -Server $vCenter | Select PowerState).PowerState -eq "PoweredOff")
	Write-Host "[SUC] '$VMName' shutdown is done !" -ForegroundColor Green
	$Line = (Get-Date -format G) + " " + "[SUC] '$VMName' shutdown is done !"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
}


### FUNCTION: START VM
Function ALIVE-VM($VM, $vCenter){
	Start-VM -VM $VMName -Confirm:$False -RunAsync | Out-Null
	Write-Host "[WAR] '$VMName' Warning: start is processing..." -ForegroundColor Yellow
	$Line = (Get-Date -format G) + " " + "[INF] '$VMName' start is processing..."
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			
	Do {$ToolStatus = (Get-vm $VMName -Server $vCenter | Get-View).Guest.ToolsRunningStatus} Until($ToolStatus -eq "guestToolsRunning")
	Write-Host "[SUC] '$VMName' VM started !`r`n" -ForegroundColor Green
	$Line = (Get-Date -format G) + " " + "[SUC] '$VMName' start is done !" + $vbcrlf
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}


### FUNCTION: Create and append LOG file
Function LogTrace ($Message){
	$Message = (Get-Date -format G) + " " + $Message
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Message -Append }

### LOG file initialization
$dat	= Get-Date -UFormat "%Y_%m_%d_%H_%M_%S"
$vbcrlf		= "`r`n"
$ScriptName	= [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition)

$RepLog     = $PathScript + "LOG"
If (!(Test-Path $RepLog)) { New-Item -Path $PathScript -ItemType Directory -Name "LOG"}
$RepLog     = $RepLog + "\"
$FicLog     = $RepLog + $ScriptName + "-" + $dat + ".log"
$FicRes     = $RepLog + $ScriptName + "-" + $dat + "_Set_RAM_CPU.csv"
$LineSep    = "=" * 70

$Line = "## BEGIN Script modification of VM RAM and CPU capacities [BT122901] ##"
If (!(Test-Path $FicLog)) {
	$Line = (Get-Date -format G) + " " + $Line
	Out-File -FilePath $FicLog -Encoding UTF8 -InputObject $Line -Append
} Else {
	LogTrace ($Line)
}
Write-Host "LogFile initialization is done !" -ForegroundColor Green
LogTrace ($LineSep + $vbcrlf)

### Load assembly & context
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
#$ErrorActionPreference= 'SilentlyContinue'
Write-Host "Setting Execution Policy to Unrestricted" -ForegroundColor Green
Set-ExecutionPolicy Unrestricted -Force | Out-Null
### https://kb.vmware.com/selfservice/microsites/search.do?language=en_US&cmd=displayKC&externalId=2009857
Write-Host "INF: Information; WAR: Warning; ERR: Error; SUC: Success`r`n" -ForegroundColor White

### TEST CSV file exist
If (-NOT (Test-Path .\VMList.csv)) {
	### EXIT Program
	[System.Windows.Forms.MessageBox]::Show("CSV file (.\VMList.csv) is missing, program is aborted !", "Abort" ,0 , 16) | Out-Null
	Write-Host "Program is aborted !" -ForegroundColor White -BackgroundColor Red
	$Line = (Get-Date -format G) + " " + "[ERR] CSV file is missing, program aborted !"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	
	Continue }

$Reponse = [System.Windows.Forms.MessageBox]::Show("Are you sure that you want to modify RAM or/and CPU capacity on each VM presents in the input file ?", "[BT122901] Confirmation" , 4, 32)
If ($Reponse -eq "Yes") {
	$VMList = Import-CSV -Delimiter ';' .\VMList.csv
	#Get-ChildItem -Path .\VMList.csv -Recurse -Force | ForEach-object { ++$CountLine }
	$CountLine = Get-ChildItem -Path .\VMList.csv -Recurse | Get-Content | Measure-Object -Line | Select -Expand Lines
	ForEach ($Item in $VMList) {
		$VMName		= $Item.("Hostname")										### VM
		$vCenter	= $Item.("vCenter")											### vCenter
		$RAM_Target	= $Item.("RAM_Targetted")									### Type in GB
		$CPU_Target	= $Item.("CPU_Targetted")									### Type in UNIT
		#[System.Windows.Forms.MessageBox]::Show("$VMName $RAM_Target")

		### USER identification initialization
		$PathScript = ($myinvocation.MyCommand.Definition | Split-Path -parent) + "\"
		$user       = "CSP_SCRIPT_ADM"
		$fickey     = "D:\Scripts\Credentials\key.crd "
		$ficcred    = "D:\Scripts\Credentials\vmware_adm.crd"
		$key        = get-content $fickey
		$pwd        = Get-Content $ficcred | ConvertTo-SecureString -key $key
		$Credential = New-Object System.Management.Automation.PSCredential $user, $pwd

		### DEFINE outfile column header
		$LOG_Head  = "VM;vCenter;Cluster;vmTools;[T0]RAM Capacity (GB);[T1]RAM Capacity (GB);[T0]CPU Capacity (CPU Unit);[T1]CPU Capacity (CPU Unit);Details"
		Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Head -Append
		
		### CSV file CHECK
		If ([String]::IsNullOrEmpty($VMName))	{	LogTrace ("[ERR] Parameter 'Hostname' is empty. Please enter the name of the target 'VM' in the CSV file." + $vbcrlf);	Continue }
		If ([String]::IsNullOrEmpty($vCenter))	{	LogTrace ("[ERR] Parameter 'vCenter' is empty. Please enter the name of the target 'vCenter' in the CSV file." + $vbcrlf);	Continue }
		If ([String]::IsNullOrEmpty($RAM_Target) -OR (-NOT [String]($RAM_Target -as [int])))	{	LogTrace ("[ERR] Parameter 'RAM_Targetted' is empty or incorrect. Please enter a value for this parameter in the CSV file." + $vbcrlf);	Continue }
		If ([String]::IsNullOrEmpty($CPU_Target) -OR (-NOT [String]($CPU_Target -as [int])))	{	LogTrace ("[ERR] Parameter 'CPU_Targetted' is empty or incorrect. Please enter a value for this parameter in the CSV file." + $vbcrlf);	Continue }
		
		### CONNECTION to vCenter
		LogTrace ("Connection to vCenter '$vCenter'")
		Write-Host "Connection to vCenter " -NoNewLine
		Write-Host "'$vCenter'... " -ForegroundColor Yellow -NoNewLine
		Write-Host "is processing" -ForegroundColor Green
	
		$rccnx = Connect-VIServer -Server $vCenter -Protocol https -Credential $Credential -WarningAction 0
		$topCnxVcc = "0"
		If ($rccnx -ne $null) {
			If ($rccnx.Isconnected) {
				$topCnxVcc = "1"	}}

		If ($topCnxVcc -ne "1") {
			LogTrace ("[ERR] Connection KO to vCenter '$vCenter' => Script halted")
			Write-Host "[ERR] Connection KO to vCenter '$vCenter' => Script halted" -ForegroundColor White -BackgroundColor Red
			$rc += 1
			
			Continue	}
		Else {
			LogTrace ("[SUC] Connection OK to vCenter '$vCenter'" + $vbcrlf)
			Write-Host "[SUC] Connection OK to vCenter '$vCenter'" -ForegroundColor Black -BackgroundColor Green	}

		$StartTime = (Get-Date)
		$Cluster	= Get-Cluster -VM $VMName -Server $vCenter | Select -Expand Name
		
		### VMTools CHECK
		$State = Get-View -ViewType VirtualMachine -Filter @{"Name" = $VMName}
		$RAM_Value = $State.Config.Hardware.MemoryMB / 1024
		$CPU_Value = $State.Config.Hardware.NumCPU
			
		Write-Host "[WAR] $($CountLine - 1) VM remaining to process`r`n" -ForegroundColor Yellow
		$CountLine -= 1
			
		Write-Host "[INF] '$VMName' start time on $StartTime" -ForegroundColor Gray
		Write-Host "[INF] '$VMName' processing..." -ForegroundColor White
		If ($State.Runtime.PowerState -eq "PoweredOn") {	Write-Host "[INF] '$VMName' is 'Power On'" }	Else {	Write-Host "[INF] '$VMName' is 'Power Off'"	}
		Write-Host "[INF] '$VMName' is present on cluster '$Cluster'" -ForegroundColor Gray
		
		If ($State.Guest.ToolsRunningStatus -eq "guestToolsRunning") {
			Write-Host "[INF] '$VMName' vmTools are running, good news !" -ForegroundColor Green
			$Line = (Get-Date -format G) + " " + "[INF] '$VMName' VMTools are running, good news !"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		} Else {
			Write-Host "[WAR] '$VMName' vmTools are NOT running, very bad news !" -ForegroundColor Red
			$Line = (Get-Date -format G) + " " + "[ERR] '$VMName' VMTools are NOT running, very bad news !"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}		
		
		$Line = (Get-Date -format G) + " " + "[INF] '$VMName' is present on cluster '$Cluster'"
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			
		### RAM capacity CHECK
		If ($RAM_Value -eq $RAM_Target) {
			Write-Host "[INF] '$VMName' has already '$RAM_Target'GB RAM, it is OK !" -ForegroundColor Green
			$Line = (Get-Date -format G) + " " + "[INF] '$VMName' has already '$RAM_Target'GB RAM, it is OK !"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			$RAM_State = "OK"	}
		Else {
			Write-Host "[WAR] '$VMName' has NOT '$RAM_Target'GB RAM but '$RAM_Value'GB RAM" -ForegroundColor Yellow
			$Line = (Get-Date -format G) + " " + "[INF] '$VMName' has NOT '$RAM_Target'GB RAM but '$RAM_Value'GB RAM"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			$RAM_State = "NOK"	}
				   
		### CPU number CHECK
		If ($CPU_Value -eq $CPU_Target) {
			Write-Host "[INF] '$VMName' has already '$CPU_Target'CPU, it is OK !" -ForegroundColor Green
			$Line = (Get-Date -format G) + " " + "[INF] '$VMName' has already '$CPU_Target'CPU, it is OK !"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			$CPU_State = "OK"	}
		Else {
			Write-Host "[WAR] '$VMName' has NOT '$CPU_Target'CPU but '$CPU_Value'CPU" -ForegroundColor Yellow
			$Line = (Get-Date -format G) + " " + "[INF] '$VMName' has NOT '$CPU_Target'CPU but '$CPU_Value'CPU"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			$CPU_State = "NOK"	}
			
		If ($RAM_State -eq "OK" -AND $CPU_State -eq "OK") {
			### Write CSV outfile
			$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;CPU and RAM already done !"
			Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
				
			$EndTime = (Get-Date)
			$ElapsedTime = $NULL
			[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
			Write-Host "[INF] '$VMName' end time on $EndTime" -ForegroundColor Gray
			Write-Host "[INF] '$VMName' time elapsed $ElapsedTime secondes`r`n" -ForegroundColor DarkGray
				
			Continue	}


		### CASE n°3 --- TEST if operation consists to decrease RAM or CPU capacity
		If ($RAM_Value -gt $RAM_Target -OR $CPU_Value -gt $CPU_Target) {
			Write-Host "[INF] '$VMName' A decrease capacity operation is processing, VM will be shutdown" -ForegroundColor White
			$Line = (Get-Date -format G) + " " + "[INF] '$VMName' A decrease capacity operation is processing, VM must be halted else it will be shutdown"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			If ($State.Runtime.PowerState -eq "PoweredOn" -AND $State.Guest.ToolsRunningStatus -eq "guestToolsRunning") {
				$VM_State = "PoweredOn"
				
				### VM is already started, shutdown VM
				Write-Host "[INF] '$VMName' VM is started and vmTools are running, shutdown is processing" -ForegroundColor White
				$Line = (Get-Date -format G) + " " + "[INF] '$VMName' VM is started, shutdown is processing"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				SHUTDOWN-VM -VM $VMName -vCenter $vCenter
				
				### MODIFY RAM Memory & CPU number (VM is halted)
				If ($RAM_State -eq "NOK") {	MODIFY-RAM -VM $VMName -vCenter $vCenter 	}
				If ($CPU_State -eq "NOK") {	MODIFY-CPU -VM $VMName -vCenter $vCenter 	}
			}

			ElseIf ($State.Runtime.PowerState -eq "PoweredOn" -AND $State.Guest.ToolsRunningStatus -ne "guestToolsRunning") {
				$VM_State = "PoweredOn"
				
				### VM is already started, vmTools not running
				Write-Host "[ERR] `n'$VMName' is started and vmTools are not running, softly shutdown is not possible !" -ForegroundColor Red
				$Line = (Get-Date -format G) + " " + "[ERR] `n'$VMName' is started and vmTools are not running, softly shutdown is not possible !"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				
				### Write CSV outfile
				$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;RAM or CPU NOT compliance, vmTools is not running"
				Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
				
				Continue
			}

			ElseIf ($State.Runtime.PowerState -eq "PoweredOff") {
				$VM_State = "PoweredOff"
		
				### VM is already halted
				Write-Host "[INF] `n'$VMName' is already shutdown !" -ForegroundColor White
				$Line = (Get-Date -format G) + " " + "[INF] '$VMName' is already shutdown !"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append

				### MODIFY RAM Memory & CPU number (VM is halted)
				If ($RAM_State -eq "NOK") {	MODIFY-RAM -VM $VMName -vCenter $vCenter 	}
				If ($CPU_State -eq "NOK") {	MODIFY-CPU -VM $VMName -vCenter $vCenter	}
			}
			
			### MODIFICATION RAM or CPU capacity CHECK
			If ($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") {
				Write-Host "[ERR] '$VMName' Error during process. RAM or CPU operation modification has failed" -ForegroundColor Red
				$Line = (Get-Date -format G) + " " + "[ERR] '$VMName' Error during process. RAM or CPU operation modification has failed"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				
				### START VM if started at initial time
				If ($VM_State -eq "PoweredOn") {	ALIVE-VM -VM $VMName -vCenter $vCenter	}
				
				### Write CSV outfile
				$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;RAM or CPU NOT compliance"
				Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
				
				Continue
			}
			
			### ADD VM notes and TAG
			If ($RAM_State -eq "NOK") {	MODIFY-RAM_INFOS -VM $VMName -vCenter $vCenter 	}
			If ($CPU_State -eq "NOK") {	MODIFY-CPU_INFOS -VM $VMName -vCenter $vCenter 	}
			
			### ADD Hot CPU or/and Hot RAM IF possible
			Enable-MemHotAdd -VM $VMName -vCenter $vCenter
			Enable-vCPUHotAdd -VM $VMName -vCenter $vCenter

			### START VM if started at initial time
			If ($VM_State -eq "PoweredOn") {	ALIVE-VM -VM $VMName -vCenter $vCenter	}
			
			### Write CSV outfile
			$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;Operation done !"
			Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
			
			Continue
		} ### END CASE n°3 ###


		### CASE n°1 --- TEST If it is an increase RAM or CPU capacity
		If ($State.Runtime.PowerState -eq "PoweredOn") {
			$VM_State = "PoweredOn"
					
			### ADD Hot CPU or/and Hot RAM IF possible
			$HotPlugStatus = (Get-VM $VMName -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select MemoryHotAddEnabled, CpuHotAddEnabled
			### TEST if HotPlug is ENABLED and if an increase operation
			If ($HotPlugStatus.MemoryHotAddEnabled -eq $True -AND $HotPlugStatus.CpuHotAddEnabled -eq $True) {
				Write-Host "[INF] '$VMName' hot add RAM and CPU parameter is enabled AND it is an increase capacity operation [HOT ADD]" -ForegroundColor White
				$Line = (Get-Date -format G) + " " + "[INF] '$VMName' hot add RAM and CPU parameter is enabled and it is an increase capacity operation [HOT ADD]"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				If ($RAM_State -eq "NOK") {	MODIFY-RAM -VM $VMName -vCenter $vCenter	}
				If ($CPU_State -eq "NOK") {	MODIFY-CPU -VM $VMName -vCenter $vCenter	}
				
				### MODIFICATION RAM or CPU capacity CHECK
				If (($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") -AND $State.Guest.ToolsRunningStatus -ne "guestToolsRunning") {
					Write-Host "[ERR] '$VMName' Error during process. RAM or CPU operation modification has failed and vmTools are not running" -ForegroundColor Red
					$Line = (Get-Date -format G) + " " + "[ERR] '$VMName' Error during process. RAM or CPU operation modification has failed and vmTools are not running"
					Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
					
					### Write CSV outfile
					$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;RAM or CPU NOT compliance, vmTools is not running"
					Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
					
					Continue
				}
				ElseIf (($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") -AND $State.Guest.ToolsRunningStatus -eq "guestToolsRunning") {
					$VM_State = "PoweredOn"
				
					### VM is already started, shutdown VM
					Write-Host "[INF] '$VMName' VM is started and vmTools are running, shutdown is processing" -ForegroundColor White
					$Line = (Get-Date -format G) + " " + "[INF] '$VMName' VM is started, shutdown is processing"
					Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
					SHUTDOWN-VM -VM $VMName -vCenter $vCenter
					
					### MODIFY RAM Memory & CPU number (VM is halted)
					If ($RAM_State -eq "NOK") {	MODIFY-RAM -VM $VMName -vCenter $vCenter 	}
					If ($CPU_State -eq "NOK") {	MODIFY-CPU -VM $VMName -vCenter $vCenter 	}
					
					### MODIFICATION RAM or CPU capacity CHECK
					If ($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") {
						Write-Host "[ERR] '$VMName' Error during process. RAM or CPU operation [COLD ADD] modification has failed" -ForegroundColor Red
						$Line = (Get-Date -format G) + " " + "[ERR] '$VMName' Error during process. RAM or CPU operation [COLD ADD] modification has failed"
						Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
						
						### START VM if started at initial time
						If ($VM_State -eq "PoweredOn") {	ALIVE-VM -VM $VMName -vCenter $vCenter	}
						
						### Write CSV outfile
						$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;RAM or CPU NOT compliance"
						Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
						
						Continue
					}
					
					### ADD VM notes and TAG
					If ($RAM_State -eq "NOK") {	MODIFY-RAM_INFOS -VM $VMName -vCenter $vCenter 	}
					If ($CPU_State -eq "NOK") {	MODIFY-CPU_INFOS -VM $VMName -vCenter $vCenter 	}
					
					### START VM if started at initial time
					If ($VM_State -eq "PoweredOn") {	ALIVE-VM -VM $VMName -vCenter $vCenter	}
					
					### Write CSV outfile
					$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;Operation done !"
					Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
					
					Continue
				}
				
				### ADD VM notes and TAG
				If ($RAM_State -eq "NOK") {	MODIFY-RAM_INFOS -VM $VMName -vCenter $vCenter 	}
				If ($CPU_State -eq "NOK") {	MODIFY-CPU_INFOS -VM $VMName -vCenter $vCenter 	}
						
				### Write CSV outfile
				$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;Operation done !"
				Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
						
				Continue
			}


			### CASE n°1.1 --- HOTPLUG DISABLED AND VM STARTED
			$VM_State = "PoweredOn"
			
			If ($State.Guest.ToolsRunningStatus -eq "guestToolsRunning") {
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $True -AND $HotPlugStatus.CpuHotAddEnabled -eq $True) {	Write-Host "[INF] '$VMName' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation [COLD SHRINK]. vmTools are running" -ForegroundColor White;	$Line = (Get-Date -format G) + " " + "[INF] '$VMName' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation. vmTools are running";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $True -AND $HotPlugStatus.CpuHotAddEnabled -eq $False) {	Write-Host "[INF] '$VMName' hot add RAM is enabled and CPU parameter is disabled but it is an increase RAM/CPU operation [COLD SHRINK]. vmTools are running" -ForegroundColor White;	$Line = (Get-Date -format G) + " " + "[INF] '$VMName' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation. vmTools are running";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $False -AND $HotPlugStatus.CpuHotAddEnabled -eq $True) {	Write-Host "[INF] '$VMName' hot add RAM is disabled and CPU parameter is enabled but it is an increase RAM/CPU operation [COLD SHRINK]. vmTools are running" -ForegroundColor White;	$Line = (Get-Date -format G) + " " + "[INF] '$VMName' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation. vmTools are running";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $False -AND $HotPlugStatus.CpuHotAddEnabled -eq $False) {	Write-Host "[INF] '$VMName' hot add RAM and CPU parameter is disabled AND it is an increase RAM/CPU operation [COLD SHRINK]. vmTools are running" -ForegroundColor White;	$Line = (Get-Date -format G) + " " + "[INF] '$VMName' hot add RAM and CPU parameter is disabled AND it is an increase RAM/CPU operation. vmTools are running";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
						
				### VM is already started, shutdown VM
				SHUTDOWN-VM -VM $VMName -vCenter $vCenter

				### MODIFY RAM Memory & CPU number (VM is halted)
				If ($RAM_State -eq "NOK") {	MODIFY-RAM -VM $VMName -vCenter $vCenter	}
				If ($CPU_State -eq "NOK") {	MODIFY-CPU -VM $VMName -vCenter $vCenter	}
				
				### MODIFICATION RAM or CPU capacity CHECK
				If ($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") {
					Write-Host "[ERR] '$VMName' Error during process. RAM or CPU operation modification has failed" -ForegroundColor Red
					$Line = (Get-Date -format G) + " " + "[ERR] '$VMName' Error during process. RAM or CPU operation modification has failed"
					Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
					
					### Write CSV outfile
					$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;RAM or CPU NOT compliance"
					Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
					
					Continue
				}
				
				### ADD VM notes and TAG
				If ($RAM_State -eq "NOK") {	MODIFY-RAM_INFOS -VM $VMName -vCenter $vCenter 	}
				If ($CPU_State -eq "NOK") {	MODIFY-CPU_INFOS -VM $VMName -vCenter $vCenter 	}

				If ($HotPlugStatus.MemoryHotAddEnabled -eq $False) {
					### ENABLE HOT add RAM
					Enable-MemHotAdd -VM $VMName -vCenter $vCenter	}

				If ($HotPlugStatus.CpuHotAddEnabled -eq $False) {
					### ENABLE HOT add CPU
					Enable-vCPUHotAdd -VM $VMName -vCenter $vCenter	}
							
				### START VM if started at initial time
				If ($VM_State -eq "PoweredOn") {	ALIVE-VM -VM $VMName -vCenter $vCenter	}
				### END CASE n°1 ###
			}
			Else {
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $True -AND $HotPlugStatus.CpuHotAddEnabled -eq $True) {	Write-Host "[INF] '$VMName' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format G) + " " + "[INF] '$VMName' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $True -AND $HotPlugStatus.CpuHotAddEnabled -eq $False) {	Write-Host "[INF] '$VMName' hot add RAM is enabled and CPU parameter is disabled but it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format G) + " " + "[INF] '$VMName' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $False -AND $HotPlugStatus.CpuHotAddEnabled -eq $True) {	Write-Host "[INF] '$VMName' hot add RAM is disabled and CPU parameter is enabled but it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format G) + " " + "[INF] '$VMName' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $False -AND $HotPlugStatus.CpuHotAddEnabled -eq $False) {	Write-Host "[INF] '$VMName' hot add RAM and CPU parameter is disabled AND it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format G) + " " + "[INF] '$VMName' hot add RAM and CPU parameter is disabled AND it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
			
				Write-Host "[ERR] '$VMName' Error during process. Shutdown, RAM or CPU operation modification has failed because vmTools are not running" -ForegroundColor Red
				$Line = (Get-Date -format G) + " " + "[ERR] '$VMName' Error during process. Shutdown, RAM or CPU operation modification has failed because vmTools are not running"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				
				### Write CSV outfile
				$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;RAM or CPU NOT compliance"
				Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
				
				Continue
			}
		}


		### CASE N°2 --- TEST If VM is halted
		Else {
			$VM_State = "PoweredOff"
				
			### VM is already halted
			Write-Host "[INF] `n'$VMName' is already shutdown !" -ForegroundColor White
			$Line = (Get-Date -format G) + " " + "[INF] '$VMName' is already shutdown !"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append

			### MODIFY RAM Memory & CPU number (VM is halted)
			If ($RAM_State -eq "NOK") {	MODIFY-RAM -VM $VMName -vCenter $vCenter	}
			If ($CPU_State -eq "NOK") {	MODIFY-CPU -VM $VMName -vCenter $vCenter	}
				
			### MODIFICATION RAM or CPU capacity CHECK
			If ($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") {
				Write-Host "[ERR] '$VMName' Error during process. RAM or CPU operation modification has failed" -ForegroundColor Red
				$Line = (Get-Date -format G) + " " + "[ERR] '$VMName' Error during process. RAM or CPU operation modification has failed"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
					
				### Write CSV outfile
				$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;RAM or CPU NOT compliance"
				Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
					
				Continue
			}
				
			### ADD VM notes and TAG
			If ($RAM_State -eq "NOK") {	MODIFY-RAM_INFOS -VM $VMName -vCenter $vCenter 	}
			If ($CPU_State -eq "NOK") {	MODIFY-CPU_INFOS -VM $VMName -vCenter $vCenter 	}
						
			### ENABLE HOT add RAM and CPU
			Enable-MemHotAdd -VM $VMName -vCenter $vCenter
			Enable-vCPUHotAdd -VM $VMName -vCenter $vCenter
		### END CASE N°2 ###
		}
		
		$EndTime = (Get-Date)
		$ElapsedTime = $NULL
		[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
		Write-Host "[INF] '$VMName' end time on $EndTime" -ForegroundColor White
		Write-Host "[INF] '$VMName' time elapsed $ElapsedTime secondes`r`n" -ForegroundColor White
			
		### Write CSV outfile
		$LOG_Line = "$VMName;$vCenter;$Cluster;$State.Guest.ToolsRunningStatus;$RAM_Value;$RAM_Target;$CPU_Value;$CPU_Target;vmTools NOT compliance"
		Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
	}

}
Else {
	### Exit Loop
	Write-Host "[INF] Program cancelled by user !" -ForegroundColor White -BackgroundColor Red
	$Line = (Get-Date -format G) + " " + "[INF] Program cancelled by user !"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
}

### Exit program
Write-Host "Program finished !" -ForegroundColor Green
$Line = "`r`n" + (Get-Date -format G) + " " + "[INF] Yippee! Program is finished !"
Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
Write-Host -NoNewLine "Press any key to quit the program...`r`n"
$NULL = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")