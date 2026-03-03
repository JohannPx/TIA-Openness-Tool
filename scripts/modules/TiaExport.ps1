# TiaExport.ps1 - Export DataBlocks as source files
# Ported from TIA_APP_V1\Functions\GenerationFunction.cs

function Get-BlockFileExtension {
    # Returns the correct file extension based on the block's programming language
    # STL -> .awl, SCL -> .scl, LAD/FBD/Graph -> .xml, default -> .scl
    param([object]$Block)

    try {
        $lang = $Block.ProgrammingLanguage.ToString()
        switch ($lang) {
            "STL"       { return ".awl" }
            "SCL"       { return ".scl" }
            "LAD"       { return ".xml" }
            "FBD"       { return ".xml" }
            "Graph"     { return ".xml" }
            "DB"        { return ".db"  }
            default     { return ".scl" }
        }
    } catch {
        return ".scl"
    }
}

function New-ExportFolder {
    param([string]$BasePath)

    if (-not $BasePath) {
        $BasePath = [Environment]::GetFolderPath('Desktop')
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $folderPath = Join-Path $BasePath "ExportDB_$timestamp"
    New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
    return $folderPath
}

function Invoke-DataBlockExport {
    param(
        [array]$SelectedBlocks,
        [string]$OutputFolder,
        [scriptblock]$OnProgress,
        [scriptblock]$OnLog
    )

    $result = @{
        SuccessCount = 0
        ErrorCount   = 0
        Errors       = @()
        OutputFolder = $OutputFolder
    }

    $state = Get-AppState
    $total = $SelectedBlocks.Count
    $current = 0

    foreach ($dbInfo in $SelectedBlocks) {
        $current++

        try {
            $block = $null
            $systemGroup = $null

            # Search the block in all PLCs
            foreach ($plcSoftware in $state.PlcSoftwareList) {
                $block = Find-BlockByNumber -BlockGroup $plcSoftware.BlockGroup -Number $dbInfo.Number -IsInstanceDB $dbInfo.IsInstanceDB
                if ($block) {
                    $systemGroup = $plcSoftware.ExternalSourceGroup
                    break
                }
            }

            if (-not $block) {
                $result.Errors += (T "MsgBlockNotFound") -f $dbInfo.Name
                $result.ErrorCount++
                continue
            }

            # Check IGenerateSource support
            if (-not ($block -is [Siemens.Engineering.SW.ExternalSources.IGenerateSource])) {
                $result.Errors += (T "MsgNoGenerateSource") -f $dbInfo.Name
                $result.ErrorCount++
                continue
            }

            # Build block list and generate
            $blocksToGenerate = [System.Collections.Generic.List[Siemens.Engineering.SW.ExternalSources.IGenerateSource]]::new()
            $blocksToGenerate.Add([Siemens.Engineering.SW.ExternalSources.IGenerateSource]$block)

            $ext = Get-BlockFileExtension -Block $block
            $fileName = "DB$($dbInfo.Number)_$($dbInfo.Name)$ext"
            $filePath = Join-Path $OutputFolder $fileName

            $systemGroup.GenerateSource(
                $blocksToGenerate,
                [System.IO.FileInfo]::new($filePath),
                [Siemens.Engineering.SW.ExternalSources.GenerateOptions]::WithDependencies
            )

            $result.SuccessCount++

            if ($OnProgress) { & $OnProgress $current $total $dbInfo.Name }
            if ($OnLog) { & $OnLog "$([char]0x2714) $fileName" }

        } catch {
            $result.Errors += "$($dbInfo.Name): $($_.Exception.Message)"
            $result.ErrorCount++
            if ($OnLog) { & $OnLog "$([char]0x2718) $($dbInfo.Name): $($_.Exception.Message)" }
        }
    }

    Set-AppStateValue -Key "LastExportResult" -Value $result
    return $result
}

function Get-ExportSummary {
    param([hashtable]$Result)

    $msg = (T "MsgExportDone") + "`n`n"
    $msg += "$([char]0x2714) " + ((T "MsgExportSuccess") -f $Result.SuccessCount) + "`n"

    if ($Result.ErrorCount -gt 0) {
        $msg += "$([char]0x2718) " + ((T "MsgExportErrors") -f $Result.ErrorCount) + "`n`n"
        $msg += (T "MsgExportDetail") + "`n"
        foreach ($err in $Result.Errors) {
            $msg += "- $err`n"
        }
    }

    $msg += "`n" + ((T "MsgExportFolder") -f $Result.OutputFolder)
    return $msg
}
