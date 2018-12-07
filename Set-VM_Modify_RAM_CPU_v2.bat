@Echo off

CD /D D:\Scripts\Powercli\Set-VM_Modify_RAM_CPU
SET ARG=VMList.csv

REM ################
Powershell.exe -executionpolicy ByPass -file Set-VM_Modify_RAM_CPU_nonInteractive_v2.ps1 %ARG%