@Echo off

CD /D D:\Scripts\Powercli\Get-ClusterLoad

REM ################
REM Nombre de jours à couvrir pour la récupération des statistiques : Indiquer une valeur en nombre de jours ou 0 pour une récupération instantanée
SET HISTODAYS=1

REM ################
REM Liste des vcenters : si plusieurs : à séparer par une virgule et entre doubles quotes
REM SET LISTVCENTERS=SWTTYV1VCSYA.yres.ytech
REM SET LISTVCENTERS=swmuzv1vcszc.zres.ztech
SET LISTVCENTERS="swmuzv1vcsza.zres.ztech,swmuzv1vcszb.zres.ztech,swmuzv1vcszc.zres.ztech,swmuzv1vcszd.zres.ztech"

REM ################
REM Liste des DataCenter à exclure : AUCUN si pas de filtrage, si plusieurs : à séparer par une virgule et entre doubles quotes
SET DCENTEXCLUS=AUCUN

REM Liste des DataCenter à inclure : TOUS si pas de filtrage, si plusieurs : à séparer par une virgule et entre doubles quotes
SET DCENTINCLUS=TOUS

REM ################
REM Liste des clusters à exclure : AUCUN si pas de filtrage, si plusieurs : à séparer par une virgule et entre doubles quotes
SET CLUSTEXCLUS=AUCUN

REM Liste des clusters à inclure : TOUS si pas de filtrage, si plusieurs : à séparer par une virgule et entre doubles quotes
SET CLUSTINCLUS=TOUS

REM ################
REM Liste des ESX à exclure : AUCUN si pas de filtrage, si plusieurs : à séparer par une virgule et entre doubles quotes
SET ESXEXCLUS=AUCUN

REM Liste des ESX à inclure : TOUS si pas de filtrage, si plusieurs : à séparer par une virgule et entre doubles quotes
SET ESXINCLUS=TOUS

REM ################
REM Liste des TAGS à exclure : AUCUN si pas de filtrage, si plusieurs : à séparer par une virgule et entre doubles quotes
SET TAGEXCLUS=AUCUN

REM Liste des TAG à inclure : TOUS si pas de filtrage, si plusieurs : à séparer par une virgule et entre doubles quotes
SET TAGINCLUS=TOUS

REM ################
REM Pourcentage du seuil de sécurité de chacun des Datastores
SET PCTSECUDS=20

REM ################
REM Capacité disque minimale pour la création d'une VM (Go)
SET CAPAMINIVM=60

REM ################
powershell.exe -executionpolicy ByPass -file Get-ClusterLoad_CVU_EnDev.ps1 %HISTODAYS% %LISTVCENTERS% %DCENTEXCLUS% %DCENTINCLUS% %CLUSTEXCLUS% %CLUSTINCLUS% %ESXEXCLUS% %ESXINCLUS% %TAGEXCLUS% %TAGINCLUS% %PCTSECUDS% %CAPAMINIVM%