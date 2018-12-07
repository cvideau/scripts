### ADVERTISEMENT
### (2018) Christophe VIDEAU - Version 1.0
Clear-Host


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
$FicRes     = $RepLog + $ScriptName + "_" + $dat + "_GET_ESX_TPS.csv"
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
$LOG_Head  = "ESXi;vCenter;Cluster;RAM installed;RAM consumed;RAM allocated"
Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Head

ForEach ($Item in $ClusterList) {
	$Cluster					= $Item.("Cluster")				### CLUSTER
	$vCenter					= $Item.("vCenter")				### vCenter
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
			
	### ESXi inventory in Cluster
	ForEach($ESX in (Get-VMHost -Location $Cluster | Where {$_.PowerState -eq "PoweredOn"} | Sort-Object Name)) {
		$statmemallocated = $NULL
		$statmemconsumed = Get-Stat -Entity $ESX -Realtime -stat mem.consumed.average | Measure-Object -Property value -Average | Select-Object -ExpandProperty Average
		Get-VMHost $ESX | Get-VM | ?{$_.PowerState -match 'PoweredOn'} | %{$statmemallocated=$statmemallocated+$_.MemoryGB}
		$statmeminstalled = Get-VMHost $ESX | Select MemoryTotalGB
		$statmeminstalled = $statmeminstalled.MemoryTotalGB
		$MemoryAllocated = $statmemallocated
		$MemoryConsumed = [math]::round(($statmemconsumed/1024000), 2)
		$MemoryInstalled = [math]::round($statmeminstalled, 0)
				
		### Write CSV outfile
		$LOG_Line = "$ESX;$vCenter;$Cluster;$MemoryInstalled;$MemoryConsumed;$MemoryAllocated"
		Out-File -Filepath $FicRes -Encoding UTF8 -InputObject $LOG_Line -Append
	}
}

### Exit program
Write-Host "Program finished !" -ForegroundColor Green
Out-File -Filepath $FicLog -Encoding UTF8 -InputObject "[INFO] Yippee! Program finished !" -Append