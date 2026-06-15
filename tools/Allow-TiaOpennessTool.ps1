#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Autorise TIA Openness Tool sur un poste ou Microsoft Defender le bloque a tort.

.DESCRIPTION
    L'executable TiaOpennessTool.exe est un binaire .NET non signe qui embarque et
    lance un script PowerShell. Ce profil declenche regulierement un FAUX POSITIF de
    Microsoft Defender ("Virus detecte"), surtout au telechargement. Le verdict est
    heuristique / machine-learning + cloud : il peut varier selon la version des
    definitions Defender, donc bloquer sur une VM et pas sur une autre issue du meme
    master.

    Ce script, a executer en tant qu'administrateur sur le MASTER / template de VM
    (les utilisateurs finaux n'ont pas les droits), fait deux choses :
      1. Met a jour les definitions Defender (le faux positif est souvent deja corrige
         cote Microsoft sur une definition recente).
      2. Ajoute des exclusions Defender ciblees sur le dossier d'installation de l'outil
         et sur le binaire temporaire de l'auto-update, pour TOUS les profils du poste.

    Les exclusions couvrent le premier lancement ET les mises a jour automatiques
    (l'updater telecharge dans %TEMP%\TiaOpennessTool_update.exe puis remplace l'exe
    installe dans %LOCALAPPDATA%\TiaOpennessTool).

.PARAMETER UsersRoot
    Racine des profils utilisateurs. Par defaut le lecteur systeme (ex. C:\Users).
    A adapter si Windows est installe sur un autre lecteur.

.NOTES
    - Le tout premier telechargement via le navigateur peut rester bloque (Defender
      supprime le fichier avant que l'exclusion ne s'applique). Deployer alors l'exe
      depuis un partage interne / l'integrer au master, plutot que via le navigateur.
    - Une exclusion ciblee par dossier applicatif n'est PAS une desactivation de
      Defender : le reste du poste reste protege.
    - En complement, signaler le faux positif a Microsoft (Microsoft Security
      Intelligence -> Submit a file for analysis) stabilise le verdict cote cloud.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Allow-TiaOpennessTool.ps1
#>
[CmdletBinding()]
param(
    [string]$UsersRoot = "$env:SystemDrive\Users"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Chemins a exclure (wildcard sur le profil pour couvrir tous les utilisateurs du poste).
$InstallDirPattern = Join-Path $UsersRoot '*\AppData\Local\TiaOpennessTool'
$UpdateExePattern  = Join-Path $UsersRoot '*\AppData\Local\Temp\TiaOpennessTool_update.exe'
$ExclusionPaths    = @($InstallDirPattern, $UpdateExePattern)

# 1. Mise a jour des definitions Defender (peut suffire a faire disparaitre le faux positif).
try {
    Write-Host "Mise a jour des definitions Microsoft Defender..."
    Update-MpSignature
    Write-Host "Definitions a jour."
} catch {
    Write-Warning "Mise a jour des definitions impossible ($($_.Exception.Message)). On continue."
}

# 2. Ajout des exclusions (idempotent : on n'ajoute que ce qui manque).
$existing = @((Get-MpPreference).ExclusionPath)
foreach ($path in $ExclusionPaths) {
    if ($existing -contains $path) {
        Write-Host "Deja exclu  : $path"
    } else {
        Add-MpPreference -ExclusionPath $path
        Write-Host "Exclu       : $path"
    }
}

Write-Host ""
Write-Host "Termine. TIA Openness Tool est autorise sur ce poste (installation + mises a jour auto)."
