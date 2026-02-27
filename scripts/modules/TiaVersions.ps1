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
