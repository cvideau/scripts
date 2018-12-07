### V2.0 (juin 2018) - Développeur: Christophe VIDEAU

Add-PSsnapin VMware.VimAutomation.Core
$WarningPreference = "SilentlyContinue"
Clear-Host
Write-Host "[BT122901] Copyright (2018) Christophe VIDEAU - Version 2.0`r`n" -ForegroundColor White

$global:State_MODIFY_RAM 			= $NULL
$global:State_MODIFY_CPU			= $NULL
$global:VM_FQDN						= $NULL
$global:BodyMail_Error 				= $NULL
$global:VM_Error 					= 0
$global:VM_Success 					= 0
$global:StopStartVM_TimeOut_Delay 	= 120	### Timeout delay in seconds
$StartTime_Script 					= (Get-Date)
$LeftTimeEstimated 					= $NULL


Function Start-Sleep($Seconds)	{
    $doneDT = (Get-Date).AddSeconds($seconds)
    while($doneDT -gt (Get-Date)) {
        $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
        $percent = ($seconds - $secondsLeft) / $seconds * 100
        Write-Progress -Activity "Waiting" -Status "Waiting..." -SecondsRemaining $secondsLeft -PercentComplete $percent
        [System.Threading.Thread]::Sleep(500)
    }
    Write-Progress -Activity "Waiting" -Status "Waiting..." -SecondsRemaining 0 -Completed
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

		Write-Host "[INFO] '$VM_FQDN' hot add RAM activation parameter is initiating" -ForegroundColor White
		$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select MemoryHotAddEnabled
		If ($HotPlugStatus.MemoryHotAddEnabled -eq $False) {
			Write-Host "[FAIL] '$VM_FQDN' Error during process. Hot add RAM parameter activation has failed" -ForegroundColor Red
			$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. Hot add RAM parameter activation has failed. Check VM state and process the change if possible... (cf. LOG report)"
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. Hot add RAM parameter activation has failed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		} Else {
			Write-Host "[SUCC] '$VM_FQDN' hot add RAM parameter activation has succedeed" -ForegroundColor Green
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' hot add RAM parameter activation has succedeed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		}
	}
	Else {
		Write-Host "[INFO] '$VM_FQDN' hot add RAM activation parameter is already activated, nothing to do" -ForegroundColor White
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' hot add RAM parameter is already activated, nothing to do"
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
		
		Write-Host "[INFO] '$VM_FQDN' hot add CPU activation parameter is initiating" -ForegroundColor White
		$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select CpuHotAddEnabled
		If ($HotPlugStatus.CpuHotAddEnabled -eq $False) {
			Write-Host "[FAIL] '$VM_FQDN' Error during process. Hot add CPU parameter activation has failed" -ForegroundColor Red
			$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. Hot add CPU parameter activation has failed. Check VM state and process the change if possible... (cf. LOG report)"
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. Hot add CPU parameter activation has failed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		} Else {
			Write-Host "[SUCC] '$VM_FQDN' hot add CPU parameter activation has succedeed" -ForegroundColor Green
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' hot add CPU parameter activation has succedeed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		}
	}
	Else {
		Write-Host "[INFO] '$VM_FQDN' hot add CPU activation parameter is already activated, nothing to do" -ForegroundColor White
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' hot add CPU parameter is already activated, nothing to do"
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
		
		Write-Host "[INFO] '$VM_FQDN' hot add RAM deactivation parameter is initiating" -ForegroundColor White
		$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select MemoryHotAddEnabled
		If ($HotPlugStatus.MemoryHotAddEnabled -eq $True) {
			Write-Host "[FAIL] '$VM_FQDN' Error during process. Hot add RAM parameter deactivation has failed" -ForegroundColor Red
			$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. Hot add RAM parameter deactivation has failed. Check VM state and process the change if possible... (cf. LOG report)"
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. Hot add RAM parameter deactivation has failed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		} Else {
			Write-Host "[SUCC] '$VM_FQDN' hot add RAM parameter deactivation has succedeed" -ForegroundColor Green
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' hot add RAM parameter deactivation has succedeed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		}
	}
	Else {
		Write-Host "[INFO] '$VM_FQDN' hot add RAM deactivation parameter is already deactivated, nothing to do" -ForegroundColor White
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' hot add RAM parameter is already deactivated, nothing to do"
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
		
		Write-Host "[INFO] '$VM_FQDN' hot add CPU deactivation parameter is initiating" -ForegroundColor White
		$HotPlugStatus = (Get-VM $VM -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select CpuHotAddEnabled
		If ($HotPlugStatus.CpuHotAddEnabled -eq $True) {
			Write-Host "[FAIL] '$VM_FQDN' Error during process. Hot add CPU parameter deactivation has failed" -ForegroundColor Red
			$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. Hot add CPU parameter deactivation has failed. Check VM state and process the change if possible... (cf. LOG report)"
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. Hot add CPU parameter deactivation has failed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		} Else {
			Write-Host "[SUCC] '$VM_FQDN' hot add CPU parameter deactivation has succedeed" -ForegroundColor Green
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' hot add CPU parameter deactivation has succedeed"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		}
	}
	Else {
		Write-Host "[INFO] '$VM_FQDN' hot add CPU deactivation parameter is already deactivated, nothing to do" -ForegroundColor White
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' hot add CPU parameter is already deactivated, nothing to do"
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	}
}
   

### FUNCTION: ADD RAM (at hot if possible)
Function MODIFY-RAM($VM, $vCenter){
Set-VM $VM -MemoryGB $RAM_Target -Confirm:$False | Out-Null
$State = Get-View -ViewType VirtualMachine -Filter @{"Name" = $VM}
If (($State.Config.Hardware.MemoryMB / 1024) -eq $RAM_Target) {
	Write-Host "[SUCC] '$VM_FQDN' has now '$RAM_Target'GB RAM" -ForegroundColor Green
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' has now '$RAM_Target'GB RAM"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	$State_MODIFY_RAM = "Success"	}
Else {
	Write-Host "[FAIL] '$VM_FQDN' Error during process. Doesn't perform RAM operation !" -ForegroundColor Red
	$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. Doesn't perform RAM operation ! Check VM state and process the change if possible... (cf. LOG report)"
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. Doesn't perform RAM operation !"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	$State_MODIFY_RAM = "Error"	}
}


### FUNCTION: ADD VM NOTES AND VM TAG FOR RAM
Function MODIFY-RAM_INFOS($VM, $vCenter){
	### Add Notes
	$VMNotes = Get-VM $VM -Server $vCenter | Select-Object -ExpandProperty Notes
	Set-VM $VM -Notes "$VMNotes`r[BT122901-OPTIMISATION RAM] La capacite RAM a ete modifiee le $(Get-Date) de '$RAM_Value'Go a '$RAM_Target'Go" -Confirm:$False | Out-Null
	Write-Host "[INFO] '$VM_FQDN' a note has been added to the VM properties" -ForegroundColor Green
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' a note has been added to the VM properties"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	
	If ($RAM_Value -lt $RAM_Target) {	$RAM_Tag = "VM_OPTIMIZE_ADD_RAM";	$RAM_Tag_Description = "[BT122901] Marqueur d'ajout de capacite RAM"	} Else {	$RAM_Tag = "VM_OPTIMIZE_SHRINK_RAM";	$RAM_Tag_Description = "[BT122901] Marqueur de réduction de capacite RAM"	}
	### CHECK if TAG exists
	$RAM_TAG_Exist = Get-Tag -Category "ADMINISTRATION" -Name $RAM_Tag -Server $vCenter -ErrorAction SilentlyContinue
	If (-NOT $RAM_TAG_Exist) {	New-Tag -Name $RAM_Tag -Category "ADMINISTRATION" -Description $RAM_Tag_Description | Out-Null	}
	
	### Add TAG
	$myTag = Get-Tag -Category "ADMINISTRATION" -Name $RAM_Tag -Server $vCenter | Out-Null
	Get-VM -Name $VM | New-TagAssignment -Tag $RAM_Tag | Out-Null
	Write-Host "[INFO] '$VM_FQDN' a TAG '$RAM_Tag' has been added to the VM properties" -ForegroundColor Green
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' a TAG '$RAM_Tag' has been added to the VM properties"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
}


### FUNCTION: ADD CPU (at hot if possible)
Function MODIFY-CPU($VM, $vCenter){
Set-VM $VM -NumCpu $CPU_Target -Confirm:$False | Out-Null
$State = Get-View -ViewType VirtualMachine -Filter @{"Name" = $VM}
If ($State.Config.Hardware.NumCPU -eq $CPU_Target) {
	Write-Host "[SUCC] '$VM_FQDN' has now '$CPU_Target'CPU" -ForegroundColor Green
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' has now '$CPU_Target'CPU"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	$State_MODIFY_CPU = "Success"	}
Else {
	Write-Host "[FAIL] '$VM_FQDN' Error during process. Doesn't perform CPU operation !" -ForegroundColor Red
	$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. Doesn't perform CPU operation ! Check VM state and process the change if possible... (cf. LOG report)"
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. Doesn't perform CPU operation !"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	$State_MODIFY_CPU = "Error"	}
}


### FUNCTION: ADD VM NOTES AND VM TAG FOR CPU
Function MODIFY-CPU_INFOS($VM, $vCenter){
	### Add Notes
	$VMNotes = Get-VM $VM -Server $vCenter | Select-Object -ExpandProperty Notes
	Set-VM $VM -Notes "$VMNotes`r[BT122901-OPTIMISATION CPU] Le nombre de CPU a ete modifie le $(Get-Date) de '$CPU_Value'CPU a '$CPU_Target'CPU" -Confirm:$False | Out-Null
	Write-Host "[INFO] '$VM_FQDN' a note has been added to the VM properties" -ForegroundColor Green
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' a note has been added to the VM properties"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	
	If ($CPU_Value -lt $CPU_Target) {	$CPU_Tag = "VM_OPTIMIZE_ADD_CPU";	$CPU_Tag_Description = "[BT122901] Marqueur d'ajout de capacite CPU"	} Else {	$CPU_Tag = "VM_OPTIMIZE_SHRINK_CPU";	$CPU_Tag_Description = "[BT122901] Marqueur de réduction de capacite CPU"	}
	### CHECK if TAG exists
	$RAM_TAG_Exist = Get-Tag -Category "ADMINISTRATION" -Name $CPU_Tag -Server $vCenter -ErrorAction SilentlyContinue
	If (-NOT $RAM_TAG_Exist) {	New-Tag -Name $CPU_Tag -Category "ADMINISTRATION" -Description $CPU_Tag_Description | Out-Null	}
	
	### Add TAG
	$myTag = Get-Tag -Category "ADMINISTRATION" -Name $CPU_Tag -Server $vCenter | Out-Null
	Get-VM -Name $VM | New-TagAssignment -Tag $CPU_Tag | Out-Null
	Write-Host "[INFO] '$VM_FQDN' a TAG '$CPU_Tag' has been added to the VM properties" -ForegroundColor Green
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' a TAG '$CPU_Tag' has been added to the VM properties"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
}


### FUNCTION: SHUTDOWN VM
Function SHUTDOWN-VM($VM, $vCenter){
	Write-Host "[WARN] '$VM_FQDN' will shutdown in 5 seconds... Press CTRL+C to CANCEL !" -ForegroundColor Yellow
	Start-Sleep 5
	Stop-VMGuest -VM $VMName -Confirm:$False -Server $vCenter | Out-Null
	Write-Host "[WARN] '$VM_FQDN' shutdown is processing !" -ForegroundColor Yellow
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' shutdown is processing !"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	
	$Stop_StartTime = (Get-Date)
	$Stop_TimeOut = $False
	[INT]$ElapsedTime = $NULL
	### Waiting stop VM
	Do {
		Start-Sleep -s 5
		$Stop_EndTime = (Get-Date)
		[INT]$ElapsedTime += 5
		Write-Host "[INFO] '$VM_FQDN' wait halt for " -NoNewLine
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
			Write-Host "[FAIL] '$VM_FQDN' shutdown task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]" -ForegroundColor Red
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' shutdown task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
	} Until (((Get-VM $VMName -Server $vCenter | Select PowerState).PowerState -eq "PoweredOff") -OR ($Stop_TimeOut -eq $True))
					
	$Stop_EndTime = (Get-Date)
	$ElapsedTime = [math]::round((($Stop_EndTime-$Stop_StartTime).TotalSeconds), 0)
	$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
	$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
	If ($Sec -eq "0") {$Sec = "00"}
	If ($Sec -eq "5") {$Sec = "05"}
					
	If ((Get-VM $VMName -Server $vCenter | Select PowerState).PowerState -eq "PoweredOff") {
		Write-Host "[SUCC] '$VM_FQDN' shutdown was done in " -NoNewLine
		Write-Host "$($Min)min. $($Sec)sec..." -ForegroundColor White
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' shutdown was done in $($Min)min. $($Sec)sec."
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
	Else {
		Write-Host "[FAIL] '$VM_FQDN' shutdown was take more than timeout delay. Over than $($MinTimeout)min. $($SecTimeout)sec." -ForegroundColor Red
		$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' was not have been processing due a shutdown timeout. Check VM state and process the change if possible... (cf. LOG report)"
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' shutdown was take more than timeout delay. Over than $($MinTimeout)min. $($SecTimeout)sec."
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
}


### FUNCTION: START VM
Function ALIVE-VM($VM, $vCenter){
	Start-VM -VM $VMName -Confirm:$False -RunAsync | Out-Null
	Write-Host "[WARN] '$VM_FQDN' start is processing..." -ForegroundColor Yellow
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' start is processing..."
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
	
	$Start_StartTime = (Get-Date)
	$Start_TimeOut = $False
	[INT]$ElapsedTime = $NULL
	### Waiting start VM
	Do {
		$ToolStatus = (Get-vm $VMName -Server $vCenter | Get-View).Guest.ToolsRunningStatus
		Start-Sleep -s 5
		$Start_EndTime = (Get-Date)
		[INT]$ElapsedTime += 5
		Write-Host "[INFO] '$VM_FQDN' wait start for " -NoNewLine
		$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
		$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
		If ($Sec -eq "0") {$Sec = "00"}
		If ($Sec -eq "5") {$Sec = "05"}
		Write-Host "$($Min)min. $($Sec)sec...`t" -ForegroundColor White -NoNewLine
		Write-Host "[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' wait start for $($Min)min. $($Sec)sec...`t[Max. $($MinTimeout)min. $($SecTimeout)sec.]"
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
		If ($ElapsedTime -ge $StopStartVM_TimeOut_Delay) {
			$Start_TimeOut = $True
			Write-Host "[FAIL] '$VM_FQDN' start task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]" -ForegroundColor Red
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' start task timeout, VM check needed [$($MinTimeout)min. $($SecTimeout)sec.]"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
	} Until (($ToolStatus -eq "guestToolsRunning") -OR ($Start_TimeOut -eq $True))
					
	$Start_EndTime = (Get-Date)
	$ElapsedTime = [math]::round((($Start_EndTime-$Start_StartTime).TotalSeconds), 0)
	$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
	$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
	If ($Sec -eq "0") {$Sec = "00"}
	If ($Sec -eq "5") {$Sec = "05"}
					
	If ((Get-VM $VMName -Server $vCenter | Select PowerState).PowerState -eq "PoweredOn") {
		Write-Host "[SUCC] '$VM_FQDN' start was done in " -NoNewLine
		Write-Host "$($Min)min. $($Sec)sec...`t" -ForegroundColor White
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' start was done in $($Min)min. $($Sec)sec."
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
	Else {
		Write-Host "[FAIL] '$VM_FQDN' was not have been processing due a start timeout. Over than $($MinTimeout)min. $($SecTimeout)sec." -ForegroundColor Red
		$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' was not have been processing due a start timeout. Check VM state and process the change if possible... (cf. LOG report)"
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' was not have been processing due a start timeout. Over than $($MinTimeout)min. $($SecTimeout)sec."
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
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

$Line = "## BEGIN Script modification of VM RAM and CPU capacities [BT122901] ##"
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
$LOG_Head  = "Date;VM;OS;vCenter;Cluster;vmTools;[T0]RAM Capacity (GB);[T1]RAM Capacity (GB);[T2]RAM Gap (GB);[T0]CPU Capacity (Unit);[T1]CPU Capacity (Unit);[T2]CPU Gap (Unit);Details"
Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Head -Append

$Reponse = [System.Windows.Forms.MessageBox]::Show("Are you sure that you want to modify RAM or/and CPU capacity on each VM presents in the input file ?", "[BT122901] Confirmation" , 4, 32)
If ($Reponse -eq "Yes") {

	$VMList = Import-CSV -Delimiter ';' .\$($args[0])
	$CountLine = Get-ChildItem -Path .\$($args[0]) -Recurse | Get-Content | Measure-Object -Line | Select -Expand Lines
	ForEach ($Item in $VMList) {
		$LeftTimeEstimated_Start 	= Get-Date													### DateTime each VM VM start
		$VMName						= $Item.("Hostname")										### VM
		$vCenter					= $Item.("vCenter")											### vCenter
		$RAM_Target					= $Item.("RAM_Targetted")									### Type in GB
		$CPU_Target					= $Item.("CPU_Targetted")									### Type in UNIT

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
		If ([String]::IsNullOrEmpty($RAM_Target) -OR (-NOT [String]($RAM_Target -as [int])))	{	LogTrace ("[FAIL] Parameter 'RAM_Targetted' is empty or incorrect. Please enter a value for this parameter in the CSV file." + $vbcrlf);	Continue }
		If ([String]::IsNullOrEmpty($CPU_Target) -OR (-NOT [String]($CPU_Target -as [int])))	{	LogTrace ("[FAIL] Parameter 'CPU_Targetted' is empty or incorrect. Please enter a value for this parameter in the CSV file." + $vbcrlf);	Continue }
		
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
				$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. Connection KO to vCenter '$vCenter'. Check vCenter state and process the change if possible... (cf. LOG report)"
				$rc += 1
				[INT]$VM_Error += 1
				Continue	}
			Else {
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] Connection OK to vCenter '$vCenter'"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				Write-Host "[SUCC] Connection OK to vCenter '$vCenter'" -ForegroundColor Black -BackgroundColor Green	}
		}
		
		$VM_OS		= (Get-View (Get-VM $VMName)).Guest.GuestFullName			### VM OS
		$VM_FQDN	= (Get-VM $VMName).Guest.HostName							### VM FQDN
		
		$vCenter_Previous = $vCenter

		$StartTime = (Get-Date)
		$StartTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
		$Cluster	= Get-Cluster -VM $VMName -Server $vCenter | Select -Expand Name
		
		### VMTools CHECK
		$State = Get-View -ViewType VirtualMachine -Filter @{"Name" = $VMName}
		$RAM_Value = $State.Config.Hardware.MemoryMB / 1024
		$CPU_Value = $State.Config.Hardware.NumCPU
	
		If ($LeftTimeEstimated -ne $NULL) {
			$Min = New-TimeSpan -Seconds $($LeftTimeEstimated * $($CountLine - 1)) | Select-Object -ExpandProperty Minutes
			$Sec = New-TimeSpan -Seconds $($LeftTimeEstimated * $($CountLine - 1)) | Select-Object -ExpandProperty Seconds
			If ($Sec -eq "0") {$Sec = "00"}
			Write-Host "[INFO] $($CountLine - 1) VM remaining to process, the remaining time is estimated at $($Min)min. $($Sec)sec. Next...`r`n" -ForegroundColor Yellow
		} Else { Write-Host "[INFO] $($CountLine - 1) VM remaining to process, the estimated remaining time is still unknown. Next...`r`n" -ForegroundColor Yellow }
		$CountLine -= 1
		
		If ($State.Guest.ToolsRunningStatus -ne "guestToolsRunning") { $VM_FQDN = $VMName; $VM_OS = "Unknown" }
		
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
			$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VMName'</B> Warning during process. vmTools are NOT running. Check VM state and process the change if necessary... (cf. LOG report)"
			Write-Host "[WARN] '$VM_FQDN' <FQDN unknown> vmTools are NOT running, very bad news !" -ForegroundColor Red
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' <FQDN unknown> VMTools are NOT running, very bad news !"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}		
		
		$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' is present on cluster '$Cluster'"
		Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			
		### RAM capacity CHECK
		If ($RAM_Value -eq $RAM_Target) {
			Write-Host "[INFO] '$VM_FQDN' has already '$RAM_Target'GB RAM, it is OK !" -ForegroundColor Green
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' has already '$RAM_Target'GB RAM, it is OK !"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			$RAM_State = "OK"	}
		Else {
			Write-Host "[WARN] '$VM_FQDN' has NOT '$RAM_Target'GB RAM but '$RAM_Value'GB RAM" -ForegroundColor Yellow
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' has NOT '$RAM_Target'GB RAM but '$RAM_Value'GB RAM"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			$RAM_State = "NOK"	}
				   
		### CPU number CHECK
		If ($CPU_Value -eq $CPU_Target) {
			Write-Host "[INFO] '$VM_FQDN' has already '$CPU_Target'CPU, it is OK !" -ForegroundColor Green
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' has already '$CPU_Target'CPU, it is OK !"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			$CPU_State = "OK"	}
		Else {
			Write-Host "[WARN] '$VM_FQDN' has NOT '$CPU_Target'CPU but '$CPU_Value'CPU" -ForegroundColor Yellow
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' has NOT '$CPU_Target'CPU but '$CPU_Value'CPU"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			$CPU_State = "NOK"	}
			
		If ($RAM_State -eq "OK" -AND $CPU_State -eq "OK") {
			### Write CSV outfile
			$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Target;$($RAM_Value-$RAM_Target);$CPU_Value;$CPU_Target;$($CPU_Value-$CPU_Target);Operation done ! CPU and RAM already done"
			Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
				
			$EndTime = (Get-Date)
			$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
			$ElapsedTime = $NULL
			[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)

			$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
			$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
			If ($Sec -eq "0") {$Sec = "00"}

			Write-Host "[INFO] '$VM_FQDN' action end time on $EndTime_Display" -ForegroundColor White
			Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
			
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			$VM_Success += 1
			
			$LeftTimeEstimated_Finish = Get-Date
			$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)

			Continue	}


		### CASE n°3 --- TEST if operation consists to decrease RAM or CPU capacity
		If ($RAM_Value -gt $RAM_Target -OR $CPU_Value -gt $CPU_Target) {
			Write-Host "[INFO] '$VM_FQDN' a decrease capacity operation is processing, VM will be shutdown if necessary" -ForegroundColor White
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' a decrease capacity operation is processing, VM must be halted else it will be shutdown"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			If ($State.Runtime.PowerState -eq "PoweredOn" -AND $State.Guest.ToolsRunningStatus -eq "guestToolsRunning") {
				$vmTools_Status = "Running"
				$VM_State = "PoweredOn"
				
				### VM is already started, shutdown VM
				Write-Host "[INFO] '$VM_FQDN' shutdown is processing" -ForegroundColor White
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' is started, shutdown is processing"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				SHUTDOWN-VM -VM $VMName -vCenter $vCenter
				
				### MODIFY RAM Memory & CPU number (VM is halted)
				If (($RAM_State -eq "NOK") -and ($Stop_TimeOut -ne $True)) {	MODIFY-RAM -VM $VMName -vCenter $vCenter 	}
				If (($CPU_State -eq "NOK") -and ($Stop_TimeOut -ne $True)) {	MODIFY-CPU -VM $VMName -vCenter $vCenter 	}
			}

			ElseIf ($State.Runtime.PowerState -eq "PoweredOn" -AND $State.Guest.ToolsRunningStatus -ne "guestToolsRunning") {
				$vmTools_Status = "Not Running"
				$VM_State = "PoweredOn"
				
				### VM is already started, vmTools not running
				Write-Host "[FAIL] '$VM_FQDN' is started and vmTools are not running, softly shutdown is not possible !" -ForegroundColor Red
				$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B> is started and vmTools are not running, softly shutdown is not possible !. Process the change if possible... (cf. LOG report)"
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' is started and vmTools are not running, softly shutdown is not possible !"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				
				### Write CSV outfile
				$EndTime = (Get-Date)
				$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
				$ElapsedTime = $NULL
				[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
				
				$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
				$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
				If ($Sec -eq "0") {$Sec = "00"}
				If ($Sec -eq "5") {$Sec = "05"}

				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				Write-Host "[INFO] '$VM_FQDN' ation end time on $EndTime_Display" -ForegroundColor White
				Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
				
				$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Value;0;$CPU_Value;$CPU_Value;0;Operation failed. RAM or CPU NOT compliance"
				Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
				$VM_Error += 1
				
				$LeftTimeEstimated_Finish = Get-Date
				$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
				
				Continue	}

			ElseIf ($State.Runtime.PowerState -eq "PoweredOff") {
				$VM_State = "PoweredOff"
		
				### VM is already halted
				Write-Host "[INFO] '$VM_FQDN' is already shutdown !" -ForegroundColor White
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' is already shutdown !"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append

				### MODIFY RAM Memory & CPU number (VM is halted)
				If ($RAM_State -eq "NOK") {	MODIFY-RAM -VM $VMName -vCenter $vCenter 	}
				If ($CPU_State -eq "NOK") {	MODIFY-CPU -VM $VMName -vCenter $vCenter	}
			}
			
			### MODIFICATION RAM or CPU capacity CHECK
			If ($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") {
				Write-Host "[FAIL] '$VM_FQDN' Error during process. RAM or/and CPU operation modification has failed" -ForegroundColor Red
				$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. RAM or/and CPU operation modification has failed. Check VM state and process the change if possible... (cf. LOG report)"
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. RAM or/and CPU operation modification has failed"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				
				### START VM if started at initial time
				If ($VM_State -eq "PoweredOn") {	ALIVE-VM -VM $VMName -vCenter $vCenter	}
				
				### Write CSV outfile
				$EndTime = (Get-Date)
				$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
				$ElapsedTime = $NULL
				[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
				
				$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
				$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
				If ($Sec -eq "0") {$Sec = "00"}
				If ($Sec -eq "5") {$Sec = "05"}

				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				Write-Host "[INFO] '$VM_FQDN' action end time on $EndTime_Display" -ForegroundColor White
				Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
				
				$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Value;0;$CPU_Value;$CPU_Value;0;Operation failed. RAM or CPU NOT compliance"
				Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
				$VM_Error += 1
				
				$LeftTimeEstimated_Finish = Get-Date
				$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
				
				Continue	}
			
			### ADD VM notes and TAG
			If ($RAM_State -eq "NOK") {	MODIFY-RAM_INFOS -VM $VMName -vCenter $vCenter 	}
			If ($CPU_State -eq "NOK") {	MODIFY-CPU_INFOS -VM $VMName -vCenter $vCenter 	}
			
			### ADD Hot CPU or/and Hot RAM IF possible
			Enable-MemHotAdd -VM $VMName -vCenter $vCenter
			Enable-vCPUHotAdd -VM $VMName -vCenter $vCenter

			### START VM if started at initial time
			If ($VM_State -eq "PoweredOn") {	ALIVE-VM -VM $VMName -vCenter $vCenter	}
			
			### Write CSV outfile
			$EndTime = (Get-Date)
			$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
			$ElapsedTime = $NULL
			[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
			
			$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
			$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
			If ($Sec -eq "0") {$Sec = "00"}
			If ($Sec -eq "5") {$Sec = "05"}

			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
			Write-Host "[INFO] '$VM_FQDN' action end time on $EndTime_Display" -ForegroundColor White
			Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
		
			$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Target;$($RAM_Value-$RAM_Target);$CPU_Value;$CPU_Target;$($CPU_Value-$CPU_Target);Operation done ! RAM or CPU compliance"
			Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
			$VM_Success += 1
			
			$LeftTimeEstimated_Finish = Get-Date
			$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
			
			Continue
		} ### END CASE n°3 ###


		### CASE n°1 --- TEST If it is an increase RAM or CPU capacity
		If ($State.Runtime.PowerState -eq "PoweredOn") {
			$VM_State = "PoweredOn"
					
			### ADD Hot CPU or/and Hot RAM IF possible
			$HotPlugStatus = (Get-VM $VMName -Server $vCenter | Select ExtensionData).ExtensionData.Config | Select MemoryHotAddEnabled, CpuHotAddEnabled
			### TEST if HotPlug is ENABLED and if an increase operation
			If ($HotPlugStatus.MemoryHotAddEnabled -eq $True -AND $HotPlugStatus.CpuHotAddEnabled -eq $True) {
				Write-Host "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is enabled AND it is an increase capacity operation [HOT ADD]" -ForegroundColor White
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' hot add RAM and CPU parameter is enabled and it is an increase capacity operation [HOT ADD]"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				If ($RAM_State -eq "NOK") {	MODIFY-RAM -VM $VMName -vCenter $vCenter	}
				If ($CPU_State -eq "NOK") {	MODIFY-CPU -VM $VMName -vCenter $vCenter	}
				
				### MODIFICATION RAM or CPU capacity CHECK
				If (($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") -AND $State.Guest.ToolsRunningStatus -ne "guestToolsRunning") {
					$vmTools_Status = "Not Running"
					Write-Host "[FAIL] '$VM_FQDN' Error during process. RAM or/and CPU operation modification has failed and vmTools are not running" -ForegroundColor Red
					$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. RAM or/and CPU operation modification has failed and vmTools are not running. Check VM state and process the change if possible... (cf. LOG report)"
					$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. RAM or/and CPU operation modification has failed and vmTools are not running"
					Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
					
					### Write CSV outfile
					$EndTime = (Get-Date)
					$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
					$ElapsedTime = $NULL
					[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
					
					$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
					$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
					If ($Sec -eq "0") {$Sec = "00"}
					If ($Sec -eq "5") {$Sec = "05"}

					$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
					Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
					Write-Host "[INFO] '$VM_FQDN' action end time on $EndTime_Display" -ForegroundColor White
					Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
		
					$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Value;0;$CPU_Value;$CPU_Value;0;Operation failed. RAM or CPU NOT compliance"
					Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
					$VM_Error += 1
					
					$LeftTimeEstimated_Finish = Get-Date
					$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
					
					Continue
				}
				ElseIf (($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") -AND $State.Guest.ToolsRunningStatus -eq "guestToolsRunning") {
					$vmTools_Status = "Running"
					$VM_State = "PoweredOn"
				
					### VM is already started, shutdown VM
					Write-Host "[INFO] '$VM_FQDN' shutdown is processing" -ForegroundColor White
					$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' is started, shutdown is processing"
					Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
					SHUTDOWN-VM -VM $VMName -vCenter $vCenter
					
					### MODIFY RAM Memory & CPU number (VM is halted)
					If (($RAM_State -eq "NOK") -and ($Stop_TimeOut -ne $True)) {	MODIFY-RAM -VM $VMName -vCenter $vCenter 	}
					If (($CPU_State -eq "NOK") -and ($Stop_TimeOut -ne $True)) {	MODIFY-CPU -VM $VMName -vCenter $vCenter 	}
					
					### MODIFICATION RAM or CPU capacity CHECK
					If ($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") {
						Write-Host "[FAIL] '$VM_FQDN' Error during process. RAM or/and CPU operation [COLD ADD] modification has failed" -ForegroundColor Red
						$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. RAM or/and CPU operation [COLD ADD] modification has failed. Check VM state and process the change if possible... (cf. LOG report)"
						$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. RAM or/and CPU operation [COLD ADD] modification has failed"
						Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
						
						### START VM if started at initial time
						If ($VM_State -eq "PoweredOn") {	ALIVE-VM -VM $VMName -vCenter $vCenter	}
						
						### Write CSV outfile
						$EndTime = (Get-Date)
						$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
						$ElapsedTime = $NULL
						[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
						
						$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
						$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
						If ($Sec -eq "0") {$Sec = "00"}
						If ($Sec -eq "5") {$Sec = "05"}

						$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
						Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
						Write-Host "[INFO] '$VM_FQDN' action end time on $EndTime_Display" -ForegroundColor White
						Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
						
						$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Value;0;$CPU_Value;$CPU_Value;0;Operation failed. RAM or CPU NOT compliance"
						Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
						$VM_Error += 1
						
						$LeftTimeEstimated_Finish = Get-Date
						$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
						
						Continue
					}
					
					### ADD VM notes and TAG
					If ($RAM_State -eq "NOK") {	MODIFY-RAM_INFOS -VM $VMName -vCenter $vCenter 	}
					If ($CPU_State -eq "NOK") {	MODIFY-CPU_INFOS -VM $VMName -vCenter $vCenter 	}
					
					### START VM if started at initial time
					If ($VM_State -eq "PoweredOn") {	ALIVE-VM -VM $VMName -vCenter $vCenter	}
					
					### Write CSV outfile
					$EndTime = (Get-Date)
					$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
					$ElapsedTime = $NULL
					[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
					
					$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
					$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
					If ($Sec -eq "0") {$Sec = "00"}
					If ($Sec -eq "5") {$Sec = "05"}

					$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
					Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
					Write-Host "[INFO] '$VM_FQDN' action end time on $EndTime_Display" -ForegroundColor White
					Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
					
					$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Target;$($RAM_Value-$RAM_Target);$CPU_Value;$CPU_Target;$($CPU_Value-$CPU_Target);Operation done ! RAM or CPU compliance"
					Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
					$VM_Success += 1
					
					$LeftTimeEstimated_Finish = Get-Date
					$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
					
					Continue
				}
				
				### ADD VM notes and TAG
				If ($RAM_State -eq "NOK") {	MODIFY-RAM_INFOS -VM $VMName -vCenter $vCenter 	}
				If ($CPU_State -eq "NOK") {	MODIFY-CPU_INFOS -VM $VMName -vCenter $vCenter 	}
						
				### Write CSV outfile
				$EndTime = (Get-Date)
				$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
				$ElapsedTime = $NULL
				[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
				
				$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
				$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
				If ($Sec -eq "0") {$Sec = "00"}
				If ($Sec -eq "5") {$Sec = "05"}

				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				Write-Host "[INFO] '$VM_FQDN' action end time on $EndTime_Display" -ForegroundColor White
				Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
				
				$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Target;$($RAM_Value-$RAM_Target);$CPU_Value;$CPU_Target;$($CPU_Value-$CPU_Target);Operation done ! RAM or CPU compliance"
				Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
				$VM_Success += 1
				
				$LeftTimeEstimated_Finish = Get-Date
				$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
						
				Continue
			}


			### CASE n°1.1 --- HOTPLUG DISABLED AND VM STARTED
			$VM_State = "PoweredOn"
			
			If ($State.Guest.ToolsRunningStatus -eq "guestToolsRunning") {
				$vmTools_Status = "Running"
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $True -AND $HotPlugStatus.CpuHotAddEnabled -eq $True) {	Write-Host "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " " + "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $True -AND $HotPlugStatus.CpuHotAddEnabled -eq $False) {	Write-Host "[INFO] '$VM_FQDN' hot add RAM is enabled and CPU parameter is disabled but it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " " + "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $False -AND $HotPlugStatus.CpuHotAddEnabled -eq $True) {	Write-Host "[INFO] '$VM_FQDN' hot add RAM is disabled and CPU parameter is enabled but it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " " + "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $False -AND $HotPlugStatus.CpuHotAddEnabled -eq $False) {	Write-Host "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is disabled AND it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " " + "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is disabled AND it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
						
				### VM is already started, shutdown VM
				SHUTDOWN-VM -VM $VMName -vCenter $vCenter

				### MODIFY RAM Memory & CPU number (VM is halted)
				If (($RAM_State -eq "NOK") -and ($Stop_TimeOut -ne $True)) {	MODIFY-RAM -VM $VMName -vCenter $vCenter	}
				If (($CPU_State -eq "NOK") -and ($Stop_TimeOut -ne $True)) {	MODIFY-CPU -VM $VMName -vCenter $vCenter	}
				
				### MODIFICATION RAM or CPU capacity CHECK
				If ($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") {
					Write-Host "[FAIL] '$VM_FQDN' Error during process. RAM or/and CPU operation modification has failed" -ForegroundColor Red
					$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. RAM or/and CPU operation modification has failed. Check VM state and process the change if possible... (cf. LOG report)"
					$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. RAM or/and CPU operation modification has failed"
					Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
					
					$EndTime = (Get-Date)
					$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
					$ElapsedTime = $NULL
					[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
					
					$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
					$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
					If ($Sec -eq "0") {$Sec = "00"}
					If ($Sec -eq "5") {$Sec = "05"}

					$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
					Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
					Write-Host "[INFO] '$VM_FQDN' action end time on $EndTime_Display" -ForegroundColor White
					Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
					
					### Write CSV outfile
					$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Value;0;$CPU_Value;$CPU_Value;0;Operation failed. RAM or CPU NOT compliance"
					Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
					$VM_Error += 1
					
					$LeftTimeEstimated_Finish = Get-Date
					$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
					
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
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $True -AND $HotPlugStatus.CpuHotAddEnabled -eq $True) {	Write-Host "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " " + "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $True -AND $HotPlugStatus.CpuHotAddEnabled -eq $False) {	Write-Host "[INFO] '$VM_FQDN' hot add RAM is enabled and CPU parameter is disabled but it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " " + "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $False -AND $HotPlugStatus.CpuHotAddEnabled -eq $True) {	Write-Host "[INFO] '$VM_FQDN' hot add RAM is disabled and CPU parameter is enabled but it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " " + "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is enabled but it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
				If ($HotPlugStatus.MemoryHotAddEnabled -eq $False -AND $HotPlugStatus.CpuHotAddEnabled -eq $False) {	Write-Host "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is disabled AND it is an increase RAM/CPU operation [COLD SHRINK]" -ForegroundColor White;	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " " + "[INFO] '$VM_FQDN' hot add RAM and CPU parameter is disabled AND it is an increase RAM/CPU operation";	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append	}
			
				Write-Host "[FAIL] '$VM_FQDN' Error during process. Shutdown, RAM or/and CPU operation modification has failed because vmTools are not running" -ForegroundColor Red
				$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. Shutdown, RAM or/and CPU operation modification has failed because vmTools are not running. Check VM state and process the change if possible... (cf. LOG report)"
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. Shutdown, RAM or/and CPU operation modification has failed because vmTools are not running"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				
				$EndTime = (Get-Date)
				$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
				$ElapsedTime = $NULL
				[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
				
				$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
				$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
				If ($Sec -eq "0") {$Sec = "00"}
				If ($Sec -eq "5") {$Sec = "05"}

				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				Write-Host "[INFO] '$VM_FQDN' action end time on $EndTime_Display" -ForegroundColor White
				Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
				
				### Write CSV outfile
				$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Value;0;$CPU_Value;$CPU_Value;0;Operation failed. RAM or CPU NOT compliance"
				Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
				$VM_Error += 1
				
				$LeftTimeEstimated_Finish = Get-Date
				$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
				
				Continue
			}
		}


		### CASE N°2 --- TEST If VM is halted
		Else {
			$VM_State = "PoweredOff"
				
			### VM is already halted
			Write-Host "[INFO] '$VM_FQDN' is already shutdown !" -ForegroundColor White
			$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' is already shutdown !"
			Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append

			### MODIFY RAM Memory & CPU number (VM is halted)
			If ($RAM_State -eq "NOK") {	MODIFY-RAM -VM $VMName -vCenter $vCenter	}
			If ($CPU_State -eq "NOK") {	MODIFY-CPU -VM $VMName -vCenter $vCenter	}
				
			### MODIFICATION RAM or CPU capacity CHECK
			If ($State_MODIFY_RAM -eq "Error" -OR $State_MODIFY_CPU -eq "Error") {
				Write-Host "[FAIL] '$VM_FQDN' Error during process. RAM or/and CPU operation modification has failed" -ForegroundColor Red
				$BodyMail_Error = $BodyMail_Error + "<BR><B>'$VM_FQDN'</B>' Error during process. RAM or/and CPU operation modification has failed. Check VM state and process the change if possible... (cf. LOG report)"
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [FAIL] '$VM_FQDN' Error during process. RAM or/and CPU operation modification has failed"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				
				$EndTime = (Get-Date)
				$EndTime_Display = (Get-Date -format 'd MMM yyyy hh:mm:ss')
				$ElapsedTime = $NULL
				[String]$ElapsedTime = [math]::round((($EndTime-$StartTime).TotalSeconds), 0)
				
				$Min = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Minutes
				$Sec = New-TimeSpan -Seconds $ElapsedTime | Select-Object -ExpandProperty Seconds
				If ($Sec -eq "0") {$Sec = "00"}
				If ($Sec -eq "5") {$Sec = "05"}

				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] '$VM_FQDN' time elapsed $($Min)min. $($Sec)sec." + $vbcrlf
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
				Write-Host "[INFO] '$VM_FQDN' action end time on $EndTime_Display" -ForegroundColor White
				Write-Host "[INFO] '$VM_FQDN' action time elapsed $($Min)min. $($Sec)sec.`r`n" -ForegroundColor White
					
				### Write CSV outfile
				$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Value;0;$CPU_Value;$CPU_Value;0;Operation failed. RAM or CPU NOT compliance"
				Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
				$VM_Error += 1
				
				$LeftTimeEstimated_Finish = Get-Date
				$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
					
				Continue
			} Else {
				Write-Host "[SUCC] '$VM_FQDN' process success. RAM or/and CPU operation modification has succedeed" -ForegroundColor Green
				$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [SUCC] '$VM_FQDN' process success. RAM or/and CPU operation modification has succedeed"
				Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
					
				### Write CSV outfile
				$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Target;$($RAM_Value-$RAM_Target);$CPU_Value;$CPU_Target;$($CPU_Value-$CPU_Target);Operation succedeed. RAM or CPU compliance"
				Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
				$VM_Success += 1
				
				$LeftTimeEstimated_Finish = Get-Date
				$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)			
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
		$LOG_Line = "$(Get-Date -format 'dd/MM/yyyy HH:mm:ss');$VMName;$VM_OS;$vCenter;$Cluster;$vmTools_Status;$RAM_Value;$RAM_Target;$($RAM_Value-$RAM_Target);$CPU_Value;$CPU_Target;$($CPU_Value-$CPU_Target);Operation done ! RAM or CPU compliance"
		Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
		$VM_Success += 1
		
		$LeftTimeEstimated_Finish = Get-Date
		$LeftTimeEstimated = [math]::round((($LeftTimeEstimated_Finish-$LeftTimeEstimated_Start).TotalSeconds), 0)
	}

}
Else {
	### Exit Loop
	Write-Host "[INFO] Program cancelled by user !" -ForegroundColor White -BackgroundColor Red
	$Line = (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] Program cancelled by user !"
	Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
}

$EndTime_Script = (Get-Date)
[String]$ElapsedTime_Script = [math]::round((($EndTime_Script-$StartTime_Script).TotalSeconds), 0)
$Min = New-TimeSpan -Seconds $ElapsedTime_Script | Select-Object -ExpandProperty Minutes
$Sec = New-TimeSpan -Seconds $ElapsedTime_Script | Select-Object -ExpandProperty Seconds
If ($Sec -eq "0") {$Sec = "00"}
If ($Sec -eq "5") {$Sec = "05"}

### Sending email
$Sender1 = "Christophe.VIDEAU-ext@ca-ts.fr"
$Sender2 = "Eric.CONSTANTIEUX-ext@ca-ts.fr"
$Sender3 = "Pascal.LAURENCE-ext@ca-ts.fr"
$CopySender1 = "Denis.MONVERT-ext@ca-ts.fr"
$CopySender2 = "Mathieu.LATREILLE-ext@ca-ts.fr"
$CopySender3 = "MCO.Infra.OS.distribues@ca-ts.fr"
$CopySender4 = "optimisation.vm@ca-ts.fr"
$From = "BT122901 - Projet optimisation RAM/CPU <BT122901.report@ca-ts.fr>"
$Subject = "[BT122901] Compte-rendu operationnel {Optimisation RAM/CPU des infrastructures VMware}"
If ($BodyMail_Error -ne $NULL) {	$Body = "Bonjour,<BR><BR>Vous trouverez en pi&egrave;ce jointe le fichier LOG et la synth&egrave;se technique de l&rsquo;op&eacute;ration.<BR><BR><U>En bref:</U><BR>Nombre de VM en succ&egrave;s:<B><I> " + $VM_Success + "</I></B><BR>Nombre de VM en &eacute;chec(s):<B><I> " + $VM_Error + "</I></B><BR>Dur&eacute;e du traitement: <I>" + $($Min) + "min. " + $($Sec) + "sec.</I><BR><BR>----------------------" + $BodyMail_Error + "<BR>----------------------<BR><BR><BR>Cordialement.<BR>L'&eacute;quipe projet (Contact: Denis MONVERT)<BR><BR>Ce message a &eacute;t&eacute; envoy&eacute; par un automate, ne pas y r&eacute;pondre."	}
Else {	$Body = "Bonjour,<BR><BR>Vous trouverez en pi&egrave;ce jointe le fichier LOG et la synth&egrave;se technique de l&rsquo;op&eacute;ration.<BR><BR><U>En bref:</U><BR>Nombre de VM en succ&egrave;s:<B><I> " + $VM_Success + "</I></B><BR>Nombre de VM en &eacute;chec(s):<B><I> " + $VM_Error + "</I></B><BR>Dur&eacute;e du traitement: <I>" + $($Min) + "min. " + $($Sec) + "sec.</I><BR><BR>Cordialement.<BR>L'&eacute;quipe projet (Contact: Denis MONVERT)<BR><BR>Ce message a &eacute;t&eacute; envoy&eacute; par un automate, ne pas y r&eacute;pondre."	}
$Attachments = $FicLog, $FicRes
$SMTP = "muz10-e1smtp-IN-DC-INT.zres.ztech"
Send-MailMessage -To $Sender1, $Sender2, $Sender3 -CC $CopySender1, $CopySender2, $CopySender3, $CopySender4 -From $From -Subject $Subject -Body $Body -Attachments $Attachments -SmtpServer $SMTP -Priority High -BodyAsHTML

### Exit program
Write-Host "[INFO] $VM_Success VM succedeed, $VM_Error VM failed... (Completed in $($Min)min. $($Sec)sec.)" -ForegroundColor White
Write-Host "Program finished ! An email has been send with reports..." -ForegroundColor Red
$Line = "`r`n" + (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] $VM_Success VM succedeed, $VM_Error VM failed... (Completed in $($Min)min. $($Sec)sec.)`r`n" + (Get-Date -format 'd MMM yyyy hh:mm:ss') + " [INFO] Yippee! Program is finished !`r An email has been send with reports..."
Out-File -Filepath $FicLog -Encoding UTF8 -InputObject $Line -Append
Write-Host -NoNewLine "Press any key to quit the program...`r`n"
$NULL = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")