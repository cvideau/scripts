@Echo off

CD /D D:\Scripts\Powercli\Set-ESX_NTP

REM ################
REM Liste des vcenters : si plusieurs : � s�parer par une virgule et entre doubles quotes
SET LISTVCENTERS=SWTTYV1VCSYA.yres.ytech
REM SET LISTVCENTERS=swmuzv1vcszc.zres.ztech
REM SET LISTVCENTERS="swmuzv1vcsza.zres.ztech,swmuzv1vcszb.zres.ztech,swmuzv1vcszc.zres.ztech,swmuzv1vcszd.zres.ztech"

REM ################
REM Liste des DataCenter � exclure : AUCUN si pas de filtrage, si plusieurs : � s�parer par une virgule et entre doubles quotes
SET DCENTEXCLUS=AUCUN

REM Liste des DataCenter � inclure : TOUS si pas de filtrage, si plusieurs : � s�parer par une virgule et entre doubles quotes
SET DCENTINCLUS=TOUS

REM ################
REM Liste des clusters � exclure : AUCUN si pas de filtrage, si plusieurs : � s�parer par une virgule et entre doubles quotes
SET CLUSTEXCLUS=AUCUN

REM Liste des clusters � inclure : TOUS si pas de filtrage, si plusieurs : � s�parer par une virgule et entre doubles quotes
SET CLUSTINCLUS=TOUS

REM ################
REM Liste des ESX � exclure : AUCUN si pas de filtrage, si plusieurs : � s�parer par une virgule et entre doubles quotes
SET ESXEXCLUS=AUCUN

REM Liste des ESX � inclure : TOUS si pas de filtrage, si plusieurs : � s�parer par une virgule et entre doubles quotes
SET ESXINCLUS=TOUS

REM ################
REM Liste des TAGS � exclure : AUCUN si pas de filtrage, si plusieurs : � s�parer par une virgule et entre doubles quotes
SET TAGEXCLUS=AUCUN

REM Liste des TAG � inclure : TOUS si pas de filtrage, si plusieurs : � s�parer par une virgule et entre doubles quotes
SET TAGINCLUS=TOUS

REM ################
powershell.exe -executionpolicy ByPass -file Set-ESX_NTP.ps1 %LISTVCENTERS% %DCENTEXCLUS% %DCENTINCLUS% %CLUSTEXCLUS% %CLUSTINCLUS% %ESXEXCLUS% %ESXINCLUS% %TAGEXCLUS% %TAGINCLUS%