# TiaExportTable.ps1 - Export DataBlocks as CSV table (Siemens Table Echange format)
# Uses ExportToXml (SimaticML) to read DB member structure, then generates CSV

# =================== XML EXPORT ===================

function Export-BlockToXml {
    param(
        [object]$Block,
        [string]$TempFolder
    )

    $fileName = "DB$($Block.Number)_$($Block.Name).xml"
    $xmlPath = Join-Path $TempFolder $fileName

    $fileInfo = [System.IO.FileInfo]::new($xmlPath)
    $Block.Export($fileInfo, [Siemens.Engineering.ExportOptions]::WithDefaults)

    return $xmlPath
}

# =================== S7 TYPE SIZES (for offset calculation) ===================

function Get-S7TypeInfo {
    # Returns size info for S7 data types (non-optimized memory layout)
    # Returns $null for complex/unknown types (Struct, UDT, FB)
    param([string]$DataType)

    $dt = $DataType.Trim('"')

    # Check Array first: "Array[lo..hi] of Type"
    if ($dt -match '^Array\s*\[(\d+)\.\.(\d+)\]\s+of\s+(.+)$') {
        $lo = [int]$Matches[1]; $hi = [int]$Matches[2]
        $elemInfo = Get-S7TypeInfo -DataType $Matches[3]
        if (-not $elemInfo) { return $null }
        $count = $hi - $lo + 1
        if ($elemInfo.IsBit) {
            # Array of Bool: packed bits, ceil(count/8) bytes
            $bytes = [Math]::Ceiling($count / 8)
            return @{ Bytes = $bytes; IsBit = $false; WordAlign = $true }
        }
        return @{ Bytes = ($elemInfo.Bytes * $count); IsBit = $false; WordAlign = $true }
    }

    # Check String[N] / WString[N]
    if ($dt -match '^String\s*\[(\d+)\]$') {
        return @{ Bytes = (2 + [int]$Matches[1]); IsBit = $false; WordAlign = $true }
    }
    if ($dt -match '^WString\s*\[(\d+)\]$') {
        return @{ Bytes = (4 + [int]$Matches[1] * 2); IsBit = $false; WordAlign = $true }
    }

    switch ($dt) {
        "Bool"          { return @{ Bytes = 0; IsBit = $true;  WordAlign = $false } }
        "Byte"          { return @{ Bytes = 1; IsBit = $false; WordAlign = $false } }
        "Char"          { return @{ Bytes = 1; IsBit = $false; WordAlign = $false } }
        "USInt"         { return @{ Bytes = 1; IsBit = $false; WordAlign = $false } }
        "SInt"          { return @{ Bytes = 1; IsBit = $false; WordAlign = $false } }
        "Word"          { return @{ Bytes = 2; IsBit = $false; WordAlign = $true  } }
        "Int"           { return @{ Bytes = 2; IsBit = $false; WordAlign = $true  } }
        "UInt"          { return @{ Bytes = 2; IsBit = $false; WordAlign = $true  } }
        "Date"          { return @{ Bytes = 2; IsBit = $false; WordAlign = $true  } }
        "S5Time"        { return @{ Bytes = 2; IsBit = $false; WordAlign = $true  } }
        "DWord"         { return @{ Bytes = 4; IsBit = $false; WordAlign = $true  } }
        "DInt"          { return @{ Bytes = 4; IsBit = $false; WordAlign = $true  } }
        "UDInt"         { return @{ Bytes = 4; IsBit = $false; WordAlign = $true  } }
        "Real"          { return @{ Bytes = 4; IsBit = $false; WordAlign = $true  } }
        "Time"          { return @{ Bytes = 4; IsBit = $false; WordAlign = $true  } }
        "Time_Of_Day"   { return @{ Bytes = 4; IsBit = $false; WordAlign = $true  } }
        "TOD"           { return @{ Bytes = 4; IsBit = $false; WordAlign = $true  } }
        "LWord"         { return @{ Bytes = 8; IsBit = $false; WordAlign = $true  } }
        "LInt"          { return @{ Bytes = 8; IsBit = $false; WordAlign = $true  } }
        "ULInt"         { return @{ Bytes = 8; IsBit = $false; WordAlign = $true  } }
        "LReal"         { return @{ Bytes = 8; IsBit = $false; WordAlign = $true  } }
        "LTime"         { return @{ Bytes = 8; IsBit = $false; WordAlign = $true  } }
        "Date_And_Time" { return @{ Bytes = 8; IsBit = $false; WordAlign = $true  } }
        "DTL"           { return @{ Bytes = 12; IsBit = $false; WordAlign = $true } }
        "String"        { return @{ Bytes = 256; IsBit = $false; WordAlign = $true } }
        "WString"       { return @{ Bytes = 512; IsBit = $false; WordAlign = $true } }
        default         { return $null }
    }
}

function Align-OffsetState {
    # Close pending bits and align to even byte (Struct boundary)
    param([hashtable]$State)

    if ($State.Bit -gt 0) {
        $State.Byte++
        $State.Bit = 0
    }
    if ($State.Byte % 2 -ne 0) {
        $State.Byte++
    }
}

# =================== SIMATICML PARSING ===================

function Get-BlockMemoryLayout {
    param([xml]$XmlDoc)

    # MemoryLayout is in the main document area (no Interface namespace)
    # Use GetElementsByTagName which ignores namespace prefixes
    $nodes = $XmlDoc.GetElementsByTagName("MemoryLayout")
    if ($nodes.Count -gt 0) { return $nodes[0].InnerText }

    return "Unknown"
}

function Parse-SimaticMlMembers {
    param(
        [string]$XmlPath,
        [int]$DbNumber
    )

    [xml]$doc = Get-Content $XmlPath -Encoding UTF8

    $nsmgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $prefix = ""

    # The Interface namespace is declared on <Sections>, NOT on the root <Document>
    # e.g. <Sections xmlns="http://www.siemens.com/automation/Openness/SW/Interface/v5">
    $sectionsNodes = $doc.GetElementsByTagName("Sections")
    if ($sectionsNodes.Count -gt 0 -and $sectionsNodes[0].NamespaceURI) {
        $nsmgr.AddNamespace("iface", $sectionsNodes[0].NamespaceURI)
        $prefix = "iface:"
    }

    # Detect memory layout
    $layout = Get-BlockMemoryLayout -XmlDoc $doc
    $isOptimized = ($layout -eq "Optimized")

    # Find the Static section containing DB members
    $staticSection = $null
    if ($prefix) {
        $staticSection = $doc.SelectSingleNode("//iface:Section[@Name='Static']", $nsmgr)
    }
    if (-not $staticSection) {
        $staticSection = $doc.SelectSingleNode("//Section[@Name='Static']")
    }

    if (-not $staticSection) {
        return @{ Members = @(); IsOptimized = $isOptimized }
    }

    # Offset tracking: enabled for non-optimized DBs
    $offsetState = @{ Byte = 0; Bit = 0; Enabled = (-not $isOptimized) }

    $members = @(Parse-MemberNodes -ParentNode $staticSection -ParentPath "" -DbNumber $DbNumber -Nsmgr $nsmgr -Prefix $prefix -OffsetState $offsetState)

    return @{
        Members     = $members
        IsOptimized = $isOptimized
    }
}

function Parse-MemberNodes {
    param(
        [System.Xml.XmlNode]$ParentNode,
        [string]$ParentPath,
        [int]$DbNumber,
        [System.Xml.XmlNamespaceManager]$Nsmgr,
        [string]$Prefix,
        [hashtable]$OffsetState,
        [string]$SectionName = ""
    )

    $results = @()

    # Get direct child Member nodes — wrap in @() to guarantee array (StrictMode safe)
    $memberNodes = @(
        if ($Nsmgr) {
            $ParentNode.SelectNodes("${Prefix}Member", $Nsmgr)
        } else {
            $ParentNode.SelectNodes("Member")
        }
    )

    foreach ($member in $memberNodes) {
        $name = $member.GetAttribute("Name")
        $dataType = $member.GetAttribute("Datatype")
        $fullPath = if ($ParentPath) { "$ParentPath.$name" } else { $name }

        # Skip system/internal members (starting with __)
        if ($name -match '^__') { continue }

        # Get comment (multi-language text - take first available)
        $comment = ""
        $commentNodes = @(
            if ($Nsmgr) {
                $member.SelectNodes("${Prefix}Comment/${Prefix}MultiLanguageText", $Nsmgr)
            } else {
                $member.SelectNodes("Comment/MultiLanguageText")
            }
        )
        if ($commentNodes.Count -gt 0 -and $commentNodes[0]) {
            $comment = $commentNodes[0].InnerText
        }

        # Detect child structure early (needed for InOut complex type check)
        $directChildren = @(
            if ($Nsmgr) {
                $member.SelectNodes("${Prefix}Member", $Nsmgr)
            } else {
                $member.SelectNodes("Member")
            }
        )
        $allSections = @(
            if ($Nsmgr) {
                $member.SelectNodes("${Prefix}Sections/${Prefix}Section", $Nsmgr)
            } else {
                $member.SelectNodes("Sections/Section")
            }
        )

        $isComplex = ($directChildren.Count -gt 0) -or ($allSections.Count -gt 0)
        $childResults = @()
        $handledAsOpaque = $false

        # InOut complex types: stored as 6-byte ANY pointer in non-optimized layout
        if ($SectionName -eq "InOut" -and $isComplex) {
            $offset = ""
            if ($OffsetState -and $OffsetState.Enabled) {
                # Close bits and word-align
                if ($OffsetState.Bit -gt 0) { $OffsetState.Byte++; $OffsetState.Bit = 0 }
                if ($OffsetState.Byte % 2 -ne 0) { $OffsetState.Byte++ }
                $offset = "$($OffsetState.Byte)"
                $OffsetState.Byte += 6  # 6-byte ANY pointer
            }
            $results += @{
                Name     = $fullPath
                DataType = $dataType
                Offset   = $offset
                Comment  = $comment
                DbNumber = $DbNumber
            }
            $handledAsOpaque = $true
        }
        # Struct (direct child Members)
        elseif ($directChildren.Count -gt 0) {
            if ($OffsetState -and $OffsetState.Enabled) { Align-OffsetState -State $OffsetState }
            $childResults += @(Parse-MemberNodes -ParentNode $member -ParentPath $fullPath -DbNumber $DbNumber -Nsmgr $Nsmgr -Prefix $Prefix -OffsetState $OffsetState -SectionName $SectionName)
            if ($OffsetState -and $OffsetState.Enabled) { Align-OffsetState -State $OffsetState }
        }
        # Sections (FB instances / UDTs)
        elseif ($allSections.Count -gt 0) {
            $isUserType = $dataType.StartsWith('"')

            if (-not $isUserType) {
                # System FB (F_TRIG, R_TRIG, TON, TOF, TCONT_CP, etc.)
                # Skip entirely — not present in export, no offset calculation
                $handledAsOpaque = $true
            } else {
                # User FB or UDT — expand all sections
                if ($OffsetState -and $OffsetState.Enabled) { Align-OffsetState -State $OffsetState }
                foreach ($sectionNode in $allSections) {
                    $secName = $sectionNode.GetAttribute("Name")
                    $childResults += @(Parse-MemberNodes -ParentNode $sectionNode -ParentPath $fullPath -DbNumber $DbNumber -Nsmgr $Nsmgr -Prefix $Prefix -OffsetState $OffsetState -SectionName $secName)
                    # Align between sections
                    if ($OffsetState -and $OffsetState.Enabled) { Align-OffsetState -State $OffsetState }
                }
            }
        }

        if ($childResults.Count -gt 0) {
            $results += $childResults
        } elseif (-not $handledAsOpaque) {
            # Leaf member — calculate offset
            $offset = ""
            if ($OffsetState -and $OffsetState.Enabled) {
                $typeInfo = Get-S7TypeInfo -DataType $dataType
                if ($typeInfo) {
                    if ($typeInfo.IsBit) {
                        # Bool: show byte.bit format (e.g. "0.0", "0.1")
                        $offset = "$($OffsetState.Byte).$($OffsetState.Bit)"
                        $OffsetState.Bit++
                        if ($OffsetState.Bit -ge 8) {
                            $OffsetState.Byte++
                            $OffsetState.Bit = 0
                        }
                    } else {
                        # Close pending bits
                        if ($OffsetState.Bit -gt 0) {
                            $OffsetState.Byte++
                            $OffsetState.Bit = 0
                        }
                        # Word alignment for 2+ byte types
                        if ($typeInfo.WordAlign -and ($OffsetState.Byte % 2 -ne 0)) {
                            $OffsetState.Byte++
                        }
                        # Non-Bool: show byte only (e.g. "48")
                        $offset = "$($OffsetState.Byte)"
                        $OffsetState.Byte += $typeInfo.Bytes
                    }
                } else {
                    # Unknown type: disable offset tracking from here
                    $OffsetState.Enabled = $false
                }
            }

            $results += @{
                Name     = $fullPath
                DataType = $dataType
                Offset   = $offset
                Comment  = $comment
                DbNumber = $DbNumber
            }
        }
    }

    return $results
}

# =================== CSV GENERATION ===================

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

function Write-CsvHeader {
    param(
        [System.IO.StreamWriter]$Writer,
        [hashtable]$PlcInfo
    )

    $Writer.WriteLine("$(T 'RowPlcName');$($PlcInfo.Name)")
    $Writer.WriteLine("$(T 'RowIpAddress');$($PlcInfo.IpAddress)")
    $Writer.WriteLine("$(T 'RowTsap');$($PlcInfo.Tsap)")
    $Writer.WriteLine(";")
    $Writer.WriteLine("$(T 'ColTag');$(T 'ColDB');$(T 'ColOffset');$(T 'ColType');$(T 'ColDescription');$(T 'ColUnit');$(T 'ColRepere');$(T 'ColCoef')")
}

function Write-CsvMemberRow {
    param(
        [System.IO.StreamWriter]$Writer,
        [hashtable]$Member,
        [bool]$IsOptimized
    )

    $offset = if ($IsOptimized) { "" } else { $Member.Offset }
    # Escape semicolons and quotes in fields
    $comment = ($Member.Comment -replace '"', '""')
    if ($comment -match '[;"]') { $comment = "`"$comment`"" }

    $name = ($Member.Name -replace '"', '""')
    if ($name -match '[;"]') { $name = "`"$name`"" }

    $Writer.WriteLine("$name;$($Member.DbNumber);$offset;$($Member.DataType);$comment;;;")
}

# =================== MAIN EXPORT ORCHESTRATION ===================

function Invoke-TableExport {
    param(
        [array]$SelectedBlocks,
        [string]$OutputFolder,
        [scriptblock]$OnProgress
    )

    $result = @{
        SuccessCount     = 0
        ErrorCount       = 0
        OptimizedDBs     = @()
        Errors           = @()
        OutputFolder     = $OutputFolder
        Files            = @()
    }

    $state = Get-AppState
    $plcInfoList = $state.PlcDeviceInfoList

    # Create temp folder for XML exports
    $tempFolder = Join-Path $OutputFolder "_temp_xml"
    New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null

    try {
        # Group selected blocks by PlcIndex
        $blocksByPlc = @{}
        foreach ($block in $SelectedBlocks) {
            $plcIdx = $block.PlcIndex
            if (-not $blocksByPlc.ContainsKey($plcIdx)) {
                $blocksByPlc[$plcIdx] = @()
            }
            $blocksByPlc[$plcIdx] += $block
        }

        $totalBlocks = $SelectedBlocks.Count
        $currentBlock = 0

        foreach ($plcIdx in ($blocksByPlc.Keys | Sort-Object)) {
            $plcBlocks = $blocksByPlc[$plcIdx]

            # Get PLC info
            $plcInfo = $plcInfoList | Where-Object { $_.PlcIndex -eq $plcIdx } | Select-Object -First 1
            if (-not $plcInfo) {
                $plcInfo = @{ PlcIndex = $plcIdx; Name = "PLC_$plcIdx"; IpAddress = ""; Rack = 0; Slot = 2; Tsap = "3.02" }
            }

            # Create CSV file
            $safeName = $plcInfo.Name -replace '[\\/:*?"<>|]', '_'
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $csvFileName = "${safeName}_Export_${timestamp}.csv"
            $csvPath = Join-Path $OutputFolder $csvFileName

            # Open StreamWriter with UTF-8 BOM (for Excel compatibility)
            $encoding = New-Object System.Text.UTF8Encoding($true)
            $writer = New-Object System.IO.StreamWriter($csvPath, $false, $encoding)

            try {
                Write-CsvHeader -Writer $writer -PlcInfo $plcInfo

                foreach ($dbInfo in $plcBlocks) {
                    $currentBlock++

                    try {
                        # Find the actual block object
                        $block = $null
                        foreach ($plcSoftware in $state.PlcSoftwareList) {
                            $block = Find-BlockByNumber -BlockGroup $plcSoftware.BlockGroup -Number $dbInfo.Number -IsInstanceDB $dbInfo.IsInstanceDB
                            if ($block) { break }
                        }

                        if (-not $block) {
                            $result.Errors += (T "MsgBlockNotFound") -f $dbInfo.Name
                            $result.ErrorCount++
                            continue
                        }

                        # Export block to XML
                        $xmlPath = Export-BlockToXml -Block $block -TempFolder $tempFolder

                        # Parse SimaticML XML
                        $parseResult = Parse-SimaticMlMembers -XmlPath $xmlPath -DbNumber $dbInfo.Number

                        if ($parseResult.IsOptimized) {
                            $result.OptimizedDBs += "DB$($dbInfo.Number)_$($dbInfo.Name)"
                            # Skip optimized DBs — no offsets available
                            $result.SuccessCount++
                            if ($OnProgress) {
                                & $OnProgress $currentBlock $totalBlocks $dbInfo.Name
                            }
                            continue
                        }

                        # Write member rows to CSV
                        foreach ($member in $parseResult.Members) {
                            Write-CsvMemberRow -Writer $writer -Member $member -IsOptimized $false
                        }

                        # Keep XML copy for debugging (in _debug subfolder)
                        $debugFolder = Join-Path $OutputFolder "_debug_xml"
                        if (-not (Test-Path $debugFolder)) {
                            New-Item -ItemType Directory -Path $debugFolder -Force | Out-Null
                        }
                        $debugCopy = Join-Path $debugFolder (Split-Path $xmlPath -Leaf)
                        Copy-Item $xmlPath $debugCopy -Force -ErrorAction SilentlyContinue

                        if ($parseResult.Members.Count -eq 0) {
                            $result.Errors += "DB$($dbInfo.Number)_$($dbInfo.Name): 0 members parsed"
                        }

                        $result.SuccessCount++

                        if ($OnProgress) {
                            & $OnProgress $currentBlock $totalBlocks $dbInfo.Name
                        }

                    } catch {
                        $result.Errors += "$($dbInfo.Name): $($_.Exception.Message)"
                        $result.ErrorCount++
                    }
                }

                $result.Files += $csvPath
            } finally {
                $writer.Close()
            }
        }
    } finally {
        # Clean up temp XML folder
        if (Test-Path $tempFolder) {
            Remove-Item $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Set-AppStateValue -Key "LastExportResult" -Value $result
    return $result
}

function Get-ExportSummary {
    param([hashtable]$Result)

    $msg = (T "MsgExportCsvDone") + "`n`n"
    $msg += "$([char]0x2714) " + ((T "MsgExportSuccess") -f $Result.SuccessCount) + "`n"

    if ($Result.ErrorCount -gt 0) {
        $msg += "$([char]0x2718) " + ((T "MsgExportErrors") -f $Result.ErrorCount) + "`n`n"
        $msg += (T "MsgExportDetail") + "`n"
        foreach ($err in $Result.Errors) {
            $msg += "- $err`n"
        }
    }

    if ($Result.OptimizedDBs.Count -gt 0) {
        $msg += "`n" + (T "MsgOptimizedWarning") + "`n"
        foreach ($db in $Result.OptimizedDBs) {
            $msg += "- $db`n"
        }
    }

    $msg += "`n" + ((T "MsgExportFolder") -f $Result.OutputFolder)

    if ($Result.Files.Count -gt 0) {
        $msg += "`n"
        foreach ($f in $Result.Files) {
            $msg += "`n- $(Split-Path $f -Leaf)"
        }
    }

    return $msg
}
