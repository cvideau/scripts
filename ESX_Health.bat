@Echo off

CD /D D:\Scripts\Powercli\ESX_Health

REM ################
REM Liste des vcenters : si plusieurs : � s�parer par une virgule et entre doubles quotes
SET LISTVCENTERS=swmuzv1vcszd.zres.ztech

REM Liste des clusters � exclure : AUCUN si pas de filtrage, si plusieurs : � s�parer par un virgule et entre doubles quotes
SET CLUSTEXCLUS=AUCUN

REM Liste des ESX � exclure : AUCUN si pas de filtrage, si plusieurs : � s�parer par un virgule et entre doubles quotes
SET ESXEXCLUS=AUCUN

REM Liste des clusters � inclure : TOUS si pas de filtrage, si plusieurs : � s�parer par un virgule et entre doubles quotes
REM SET CLUSTINCLUS=TOUS
SET CLUSTINCLUS="CL_MU_HDI_Z12,CL_MU_HDI_Z13,CL_MU_HDM_Z12,CL_MU_HDM_Z13"

REM Liste des ESX � inclure : TOUS si pas de filtrage, si plusieurs : � s�parer par un virgule et entre doubles quotes
SET ESXINCLUS=TOUS

REM ################
Powershell.exe -executionpolicy ByPass -file ESX_Health_0.3.ps1 %LISTVCENTERS% %CLUSTEXCLUS% %ESXEXCLUS% %CLUSTINCLUS% %ESXINCLUS%