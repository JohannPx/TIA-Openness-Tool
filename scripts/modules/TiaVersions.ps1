# TiaVersions.ps1 - Multi-version TIA Portal DLL detection and loading

function Get-InstalledTiaVersions {
    $versions = @()
    $basePath = "C:\Program Files\Siemens\Automation"

    if (Test-Path $basePath) {
        Get-ChildItem -Path $basePath -Directory -Filter "Portal V*" -ErrorAction SilentlyContinue | ForEach-Object {
            $dirName = $_.Name
            if ($dirName -match 'Portal V(\d+)') {
                $vNum = $Matches[1]
                $vLabel = "V$vNum"
                $dllPath = Join-Path $_.FullName "PublicAPI\V$vNum\Siemens.Engineering.dll"
                if (Test-Path $dllPath) {
                    $versions += @{
                        Version     = $vLabel
                        MajorNumber = [int]$vNum
                        DllPath     = $dllPath
                    }
                }
            }
        }
    }

    return $versions | Sort-Object { $_.MajorNumber }
}

function Initialize-TiaOpenness {
    param([string]$DllPath)

    $state = Get-AppState

    # If DLL already loaded for a different version, warn
    if ($state.DllLoaded -and $state.DllPath -ne $DllPath) {
        return @{ Success = $false; Message = T "MsgRestartRequired" }
    }

    # If already loaded same path, skip
    if ($state.DllLoaded -and $state.DllPath -eq $DllPath) {
        return @{ Success = $true; Message = "" }
    }

    try {
        [System.Reflection.Assembly]::LoadFrom($DllPath) | Out-Null
        Set-AppStateValue -Key "DllPath" -Value $DllPath
        Set-AppStateValue -Key "DllLoaded" -Value $true
        return @{ Success = $true; Message = "" }
    } catch {
        try {
            Add-Type -LiteralPath $DllPath -ErrorAction Stop
            Set-AppStateValue -Key "DllPath" -Value $DllPath
            Set-AppStateValue -Key "DllLoaded" -Value $true
            return @{ Success = $true; Message = "" }
        } catch {
            return @{ Success = $false; Message = $_.Exception.Message }
        }
    }
}

function Get-TiaVersionFromProcess {
    param([int]$ProcessId)
    try {
        $wmiResult = Get-WmiObject -Query "SELECT ExecutablePath FROM Win32_Process WHERE ProcessId = $ProcessId" -ErrorAction SilentlyContinue
        if ($wmiResult -and $wmiResult.ExecutablePath -match 'Portal V(\d+)') {
            return "V$($Matches[1])"
        }
    } catch {}
    return $null
}

function Get-RunningTiaPortalVersions {
    # Detecte les instances TIA Portal reellement en cours d'execution, via WMI, en se
    # basant sur le chemin de l'executable (Portal Vxx). Independant de l'API Openness,
    # qui ne voit que les instances de la version de DLL chargee : sert donc a choisir
    # le bon defaut de version au demarrage et a diagnostiquer un scan vide.
    $running = @()
    try {
        $procs = Get-WmiObject -Query "SELECT ProcessId, ExecutablePath FROM Win32_Process WHERE Name = 'Siemens.Automation.Portal.exe'" -ErrorAction SilentlyContinue
        foreach ($p in $procs) {
            if ($p.ExecutablePath -match 'Portal V(\d+)') {
                $running += @{
                    ProcessId   = [int]$p.ProcessId
                    Version     = "V$($Matches[1])"
                    MajorNumber = [int]$Matches[1]
                }
            }
        }
    } catch {}
    return $running
}

function Resolve-SelectedDllPath {
    # Retourne le chemin de la DLL Openness correspondant a la version actuellement
    # selectionnee dans l'interface (sans la charger).
    $state = Get-AppState
    if (-not $state.SelectedVersion) { return $null }
    $match = $state.InstalledVersions | Where-Object { $_.Version -eq $state.SelectedVersion } | Select-Object -First 1
    if ($match) { return $match.DllPath }
    return $null
}

function Confirm-TiaDllLoaded {
    # Charge (paresseusement) la DLL Openness de la version selectionnee si ce n'est pas
    # deja fait. Le chargement est differe jusqu'au scan/connexion : tant qu'aucune DLL
    # n'est chargee, l'utilisateur peut librement changer de version. Une fois une version
    # chargee, .NET ne permet pas d'en charger une autre dans le meme processus -- il faut
    # alors redemarrer l'outil.
    if ((Get-AppState).DllLoaded) { return @{ Success = $true; Message = "" } }
    $dllPath = Resolve-SelectedDllPath
    if (-not $dllPath) { return @{ Success = $false; Message = T "MsgDllRequired" } }
    return Initialize-TiaOpenness -DllPath $dllPath
}
