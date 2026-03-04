# TiaDataBlocks.ps1 - DataBlock enumeration and filtering
# Ported from TIA_APP_V1\Functions\DataBlockFunction.cs

function Get-BlocksFromGroup {
    param(
        [object]$BlockGroup,
        [int]$PlcIndex
    )

    $blocks = @()

    foreach ($block in $BlockGroup.Blocks) {
        $isDataBlock = $block -is [Siemens.Engineering.SW.Blocks.DataBlock]
        $isInstanceDB = $block -is [Siemens.Engineering.SW.Blocks.InstanceDB]

        if ($isDataBlock -or $isInstanceDB) {
            $blocks += @{
                Name                 = $block.Name
                Number               = $block.Number
                IsInstanceDB         = $isInstanceDB
                DisplayType          = if ($isInstanceDB) { T "TypeInstance" } else { T "TypeGlobal" }
                BadgeColor           = if ($isInstanceDB) { "#F59E0B" } else { "#3B82F6" }
                BadgeBackgroundColor = if ($isInstanceDB) { "#FEF3C7" } else { "#DBEAFE" }
                PlcIndex             = $PlcIndex
                IsSelected           = $false
                Index                = 0
            }
        }
    }

    foreach ($subGroup in $BlockGroup.Groups) {
        $blocks += Get-BlocksFromGroup -BlockGroup $subGroup -PlcIndex $PlcIndex
    }

    return $blocks
}

function Get-AllDataBlocks {
    $state = Get-AppState
    if (-not $state.IsConnected) {
        throw (T "MsgConnectFirst")
    }

    $allBlocks = @()
    $plcIndex = 1

    foreach ($plcSoftware in $state.PlcSoftwareList) {
        $plcBlocks = @(Get-BlocksFromGroup -BlockGroup $plcSoftware.BlockGroup -PlcIndex $plcIndex)

        # Sort by Number within each PLC
        $sortedBlocks = $plcBlocks | Sort-Object { $_.Number }
        $allBlocks += $sortedBlocks

        $plcIndex++
    }

    # Assign global index
    $index = 1
    foreach ($block in $allBlocks) {
        $block.Index = $index++
    }

    Set-AppStateValue -Key "AllDataBlocks" -Value $allBlocks
    Apply-DataBlockFilter
    return $allBlocks
}

function Apply-DataBlockFilter {
    $state = Get-AppState
    $all = $state.AllDataBlocks

    if ($state.HideInstanceDBs) {
        $filtered = @($all | Where-Object { -not $_.IsInstanceDB })
    } else {
        $filtered = @($all)
    }

    # Re-index
    $index = 1
    foreach ($block in $filtered) {
        $block.Index = $index++
    }

    Set-AppStateValue -Key "FilteredDataBlocks" -Value $filtered
}

function Find-BlockByNumber {
    param(
        [object]$BlockGroup,
        [int]$Number,
        [bool]$IsInstanceDB
    )

    foreach ($block in $BlockGroup.Blocks) {
        if ($block.Number -eq $Number) {
            $blockIsDataBlock = $block -is [Siemens.Engineering.SW.Blocks.DataBlock]
            $blockIsInstanceDB = $block -is [Siemens.Engineering.SW.Blocks.InstanceDB]

            # Skip non-DB blocks (OBs, FBs, FCs also live in BlockGroup.Blocks)
            if (-not $blockIsDataBlock -and -not $blockIsInstanceDB) { continue }

            if ($IsInstanceDB -and $blockIsInstanceDB) {
                return $block
            }
            elseif (-not $IsInstanceDB -and $blockIsDataBlock -and -not $blockIsInstanceDB) {
                return $block
            }
        }
    }

    foreach ($subGroup in $BlockGroup.Groups) {
        $found = Find-BlockByNumber -BlockGroup $subGroup -Number $Number -IsInstanceDB $IsInstanceDB
        if ($found) { return $found }
    }

    return $null
}

function Find-BlockByName {
    param(
        [object]$BlockGroup,
        [string]$Name
    )

    foreach ($block in $BlockGroup.Blocks) {
        if ($block.Name -eq $Name) { return $block }
    }

    foreach ($subGroup in $BlockGroup.Groups) {
        $found = Find-BlockByName -BlockGroup $subGroup -Name $Name
        if ($found) { return $found }
    }

    return $null
}

function Find-TypeByName {
    # Search UDT types in PlcSoftware.TypeGroup (PlcTypeSystemGroup)
    param(
        [object]$TypeGroup,
        [string]$Name
    )

    foreach ($t in $TypeGroup.Types) {
        if ($t.Name -eq $Name) { return $t }
    }

    foreach ($subGroup in $TypeGroup.Groups) {
        $found = Find-TypeByName -TypeGroup $subGroup -Name $Name
        if ($found) { return $found }
    }

    return $null
}
