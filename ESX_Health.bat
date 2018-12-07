@Echo off

CD /D D:\Scripts\Powercli\ESX_Health\DEV

REM ################
REM Liste des vcenters : si plusieurs : à séparer par une virgule et entre doubles quotes
SET LISTVCENTERS=SWMUZV1VCSZA.zres.ztech,SWMUZV1VCSZB.zres.ztech,SWMUZV1VCSZC.zres.ztech,SWMUZV1VCSZD.zres.ztech,SWTTYV1VCSYA.yres.ytech,SWTTYV1VCSYB.yres.ytech,SWMUZV1VCSQ1.zres.ztech

REM Liste des clusters à exclure : AUCUN si pas de filtrage, si plusieurs : à séparer par un virgule et entre doubles quotes
SET CLUSTEXCLUS=AUCUN

REM Liste des ESX à exclure : AUCUN si pas de filtrage, si plusieurs : à séparer par un virgule et entre doubles quotes
SET ESXEXCLUS=AUCUN

REM Liste des clusters à inclure : TOUS si pas de filtrage, si plusieurs : à séparer par un virgule et entre doubles quotes
REM SET CLUSTINCLUS=TOUS
SET CLUSTINCLUS=TOUS

REM Liste des ESX à inclure : TOUS si pas de filtrage, si plusieurs : à séparer par un virgule et entre doubles quotes
REM SET ESXINCLUS=sxmuzhvhdich.zres.ztech,sxmuzhvhdidk.zres.ztech,sxmuzhvhdisv.zres.ztech,sxmuzhvhdiz4.zres.ztech,sxmuzhvhdizh.zres.ztech,sxmuzhvhdmch.zres.ztech,sxmuzhvhdmdk.zres.ztech,sxmuzhvhdmsv.zres.ztech,sxmuzhvhdmz4.zres.ztech,sxmuzhvhdmzh.zres.ztech,sxmuzhvhdi1n.zres.ztech,sxmuzhvhdip1.zres.ztech,sxmuzhvhdm1n.zres.ztech,sxmuzhvhdmp1.zres.ztech
SET ESXINCLUS=TOUS

REM ################
Powershell.exe -executionpolicy ByPass -file ESX_Health_0.4.ps1 %LISTVCENTERS%
REM Powershell.exe -executionpolicy ByPass -file ESX_Health_0.3.ps1 %LISTVCENTERS% %CLUSTEXCLUS% %ESXEXCLUS% %CLUSTINCLUS% %ESXINCLUS%