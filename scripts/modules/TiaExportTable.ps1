# TiaExportTable.ps1 - Export DataBlocks as CSV table (Siemens Table Echange format)
# Uses ExportToXml (SimaticML) to read DB member structure, then generates CSV

# =================== XML EXPORT ===================

function Export-BlockToXml {
    param(
        [object]$Block,
        [string]$TempFolder
    )

    # PlcType (UDT) has no .Number property — use Name only
    $blockName = $Block.Name -replace '[\\/:*?"<>|]', '_'
    try {
        $fileName = "DB$($Block.Number)_${blockName}.xml"
    } catch {
        $fileName = "TYPE_${blockName}.xml"
    }
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

function Get-CommentsFromXml {
    # Recursively extract member comments from a SimaticML XML (FB or UDT)
    # Returns a hashtable: "Path.To.Member" → "comment text"
    param(
        [System.Xml.XmlNode]$ParentNode,
        [string]$ParentPath = ""
    )

    $comments = @{}
    foreach ($child in $ParentNode.ChildNodes) {
        if ($child.LocalName -ne "Member") { continue }

        $name = $child.GetAttribute("Name")
        $fullPath = if ($ParentPath) { "$ParentPath.$name" } else { $name }

        # Extract comment
        foreach ($cn in $child.ChildNodes) {
            if ($cn.LocalName -eq "Comment") {
                foreach ($mlText in $cn.ChildNodes) {
                    if ($mlText.LocalName -eq "MultiLanguageText" -and -not [string]::IsNullOrEmpty($mlText.InnerText)) {
                        $comments[$fullPath] = $mlText.InnerText
                        break
                    }
                }
                break
            }
        }

        # Recurse into child Members (Struct)
        $hasChildMembers = $false
        foreach ($sub in $child.ChildNodes) {
            if ($sub.LocalName -eq "Member") { $hasChildMembers = $true; break }
        }
        if ($hasChildMembers) {
            $childComments = Get-CommentsFromXml -ParentNode $child -ParentPath $fullPath
            foreach ($key in $childComments.Keys) { $comments[$key] = $childComments[$key] }
        }

        # Recurse into Sections (UDT/FB sub-types)
        foreach ($sub in $child.ChildNodes) {
            if ($sub.LocalName -eq "Sections") {
                foreach ($section in $sub.ChildNodes) {
                    if ($section.LocalName -eq "Section") {
                        $childComments = Get-CommentsFromXml -ParentNode $section -ParentPath $fullPath
                        foreach ($key in $childComments.Keys) { $comments[$key] = $childComments[$key] }
                    }
                }
            }
        }
    }
    return $comments
}

function Get-InstanceOfName {
    # Detect if a SimaticML XML is an instance DB and return the FB/UDT name
    param([xml]$XmlDoc)

    $node = $XmlDoc.GetElementsByTagName("InstanceOfName")
    if ($node.Count -gt 0 -and -not [string]::IsNullOrEmpty($node[0].InnerText)) {
        return $node[0].InnerText
    }
    return $null
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
        [string]$SectionName = "",
        [string]$CurrentUdtType = "",
        [string]$UdtRelativePath = ""
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

        # Get comment (multi-language text - take first non-empty)
        $comment = ""
        foreach ($childNode in $member.ChildNodes) {
            if ($childNode.LocalName -eq "Comment") {
                foreach ($mlText in $childNode.ChildNodes) {
                    if ($mlText.LocalName -eq "MultiLanguageText" -and -not [string]::IsNullOrEmpty($mlText.InnerText)) {
                        $comment = $mlText.InnerText
                        break
                    }
                }
                break
            }
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
        # Not shown in CSV but offset must be accounted for
        if ($SectionName -eq "InOut" -and $isComplex) {
            if ($OffsetState -and $OffsetState.Enabled) {
                # Close bits and word-align
                if ($OffsetState.Bit -gt 0) { $OffsetState.Byte++; $OffsetState.Bit = 0 }
                if ($OffsetState.Byte % 2 -ne 0) { $OffsetState.Byte++ }
                $OffsetState.Byte += 6  # 6-byte ANY pointer
            }
            $handledAsOpaque = $true
        }
        # Struct (direct child Members) — keep same UDT context, extend relative path
        elseif ($directChildren.Count -gt 0) {
            if ($OffsetState -and $OffsetState.Enabled) { Align-OffsetState -State $OffsetState }
            $childUdtPath = if ($CurrentUdtType) {
                if ($UdtRelativePath) { "$UdtRelativePath.$name" } else { $name }
            } else { "" }
            $childResults += @(Parse-MemberNodes -ParentNode $member -ParentPath $fullPath -DbNumber $DbNumber -Nsmgr $Nsmgr -Prefix $Prefix -OffsetState $OffsetState -SectionName $SectionName -CurrentUdtType $CurrentUdtType -UdtRelativePath $childUdtPath)
            if ($OffsetState -and $OffsetState.Enabled) { Align-OffsetState -State $OffsetState }
        }
        # Sections (FB instances / UDTs) — enter new UDT context
        elseif ($allSections.Count -gt 0) {
            $isUserType = $dataType.StartsWith('"')

            if (-not $isUserType) {
                # System FB (F_TRIG, R_TRIG, TON, TOF, TCONT_CP, etc.)
                # Skip entirely — not present in export, no offset calculation
                $handledAsOpaque = $true
            } else {
                # User FB or UDT — expand all sections, reset UDT context to this type
                $udtName = $dataType.Trim('"')
                if ($OffsetState -and $OffsetState.Enabled) { Align-OffsetState -State $OffsetState }
                foreach ($sectionNode in $allSections) {
                    $secName = $sectionNode.GetAttribute("Name")
                    $childResults += @(Parse-MemberNodes -ParentNode $sectionNode -ParentPath $fullPath -DbNumber $DbNumber -Nsmgr $Nsmgr -Prefix $Prefix -OffsetState $OffsetState -SectionName $secName -CurrentUdtType $udtName -UdtRelativePath "")
                    # Align between sections
                    if ($OffsetState -and $OffsetState.Enabled) { Align-OffsetState -State $OffsetState }
                }
            }
        }

        if ($childResults.Count -gt 0) {
            $results += $childResults
        } elseif (-not $handledAsOpaque) {
            # Compute UDT-relative path for comment resolution
            $leafUdtPath = if ($CurrentUdtType) {
                if ($UdtRelativePath) { "$UdtRelativePath.$name" } else { $name }
            } else { $fullPath }

            # Check if this is an Array type to expand into individual elements
            $dtClean = $dataType.Trim('"')
            $isArray = $dtClean -match '^Array\s*\[(\d+)\.\.(\d+)\]\s+of\s+(.+)$'

            if ($isArray) {
                $arrLo = [int]$Matches[1]; $arrHi = [int]$Matches[2]
                $baseType = $Matches[3].Trim()
                $elemInfo = Get-S7TypeInfo -DataType $baseType

                if ($elemInfo) {
                    # Word-align before array start
                    if ($OffsetState -and $OffsetState.Enabled) {
                        if ($OffsetState.Bit -gt 0) { $OffsetState.Byte++; $OffsetState.Bit = 0 }
                        if ($OffsetState.Byte % 2 -ne 0) { $OffsetState.Byte++ }
                    }

                    for ($arrIdx = $arrLo; $arrIdx -le $arrHi; $arrIdx++) {
                        $elemOffset = ""
                        if ($OffsetState -and $OffsetState.Enabled) {
                            if ($elemInfo.IsBit) {
                                $elemOffset = "$($OffsetState.Byte).$($OffsetState.Bit)"
                                $OffsetState.Bit++
                                if ($OffsetState.Bit -ge 8) { $OffsetState.Byte++; $OffsetState.Bit = 0 }
                            } else {
                                $elemOffset = "$($OffsetState.Byte).0"
                                $OffsetState.Byte += $elemInfo.Bytes
                            }
                        }
                        $results += @{
                            Name     = "${fullPath}[${arrIdx}]"
                            DataType = $baseType
                            Offset   = $elemOffset
                            Comment  = $comment
                            DbNumber = $DbNumber
                            UdtType  = $CurrentUdtType
                            UdtPath  = $leafUdtPath
                        }
                    }
                } else {
                    # Unknown base type — emit as single entry without expansion
                    $results += @{
                        Name     = $fullPath
                        DataType = $dataType
                        Offset   = ""
                        Comment  = $comment
                        DbNumber = $DbNumber
                        UdtType  = $CurrentUdtType
                        UdtPath  = $leafUdtPath
                    }
                    if ($OffsetState -and $OffsetState.Enabled) { $OffsetState.Enabled = $false }
                }
            } else {
                # Regular leaf member — calculate offset
                $offset = ""
                if ($OffsetState -and $OffsetState.Enabled) {
                    $typeInfo = Get-S7TypeInfo -DataType $dataType
                    if ($typeInfo) {
                        if ($typeInfo.IsBit) {
                            $offset = "$($OffsetState.Byte).$($OffsetState.Bit)"
                            $OffsetState.Bit++
                            if ($OffsetState.Bit -ge 8) { $OffsetState.Byte++; $OffsetState.Bit = 0 }
                        } else {
                            if ($OffsetState.Bit -gt 0) { $OffsetState.Byte++; $OffsetState.Bit = 0 }
                            if ($typeInfo.WordAlign -and ($OffsetState.Byte % 2 -ne 0)) { $OffsetState.Byte++ }
                            $offset = "$($OffsetState.Byte).0"
                            $OffsetState.Byte += $typeInfo.Bytes
                        }
                    } else {
                        $OffsetState.Enabled = $false
                    }
                }

                $results += @{
                    Name     = $fullPath
                    DataType = $dataType
                    Offset   = $offset
                    Comment  = $comment
                    DbNumber = $DbNumber
                    UdtType  = $CurrentUdtType
                    UdtPath  = $leafUdtPath
                }
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
    if (-not (Test-Path $BasePath)) {
        New-Item -ItemType Directory -Path $BasePath -Force | Out-Null
    }
    return $BasePath
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

# =================== EWON VAR_LST FORMAT ===================

# S7 address format suffixes for Ewon (from edgeMap S7_FORMAT)
$Script:S7_EWON_FORMAT = @{
    "Bool"="B"; "Byte"="B"; "SInt"="B"; "USInt"="B"
    "Int"="S"; "UInt"="W"; "Word"="W"; "S5Time"="S"; "Date"="S"
    "DInt"="L"; "UDInt"="D"; "DWord"="D"; "Real"="F"; "Time"="L"; "TOD"="L"; "Time_Of_Day"="L"
    "LWord"="D"; "LInt"="L"; "ULInt"="D"; "LReal"="F"; "LTime"="L"
    "Date_And_Time"="D"; "DTL"="D"; "String"="W"; "WString"="W"
}

# Ewon display type: 0=BOOL, 1=Float, 2=Int(8/16bit), 3=DWord(32/64bit)
$Script:EWON_TYPE_MAP = @{
    "Bool"=0
    "Byte"=2; "SInt"=2; "USInt"=2; "Char"=2
    "Int"=2; "UInt"=2; "Word"=2; "S5Time"=2; "Date"=2
    "DInt"=3; "UDInt"=3; "DWord"=3; "Time"=3; "TOD"=3; "Time_Of_Day"=3
    "Real"=1; "LReal"=1
    "LInt"=3; "ULInt"=3; "LWord"=3; "LTime"=3
    "Date_And_Time"=3; "DTL"=3; "String"=2; "WString"=2
}

# 62-column var_lst header (Ewon Flexy format)
$Script:VAR_LST_COLUMNS = @(
    @{N="Id";T="num"};@{N="Name";T="str"};@{N="Description";T="str"};@{N="ServerName";T="str"}
    @{N="TopicName";T="str"};@{N="Address";T="str"};@{N="Coef";T="num"};@{N="Offset";T="num"}
    @{N="LogEnabled";T="num"};@{N="AlEnabled";T="num"};@{N="AlBool";T="num"};@{N="MemTag";T="num"}
    @{N="MbsTcpEnabled";T="num"};@{N="MbsTcpFloat";T="num"};@{N="SnmpEnabled";T="num"}
    @{N="RTLogEnabled";T="num"};@{N="AlAutoAck";T="num"};@{N="ForceRO";T="num"}
    @{N="SnmpOID";T="num"};@{N="AutoType";T="num"};@{N="AlHint";T="str"};@{N="AlHigh";T="num"}
    @{N="AlLow";T="num"};@{N="AlTimeDB";T="num"};@{N="AlLevelDB";T="num"}
    @{N="IVGroupA";T="num"};@{N="IVGroupB";T="num"};@{N="IVGroupC";T="num"};@{N="IVGroupD";T="num"}
    @{N="PageId";T="num"};@{N="RTLogWindow";T="num"};@{N="RTLogTimer";T="num"}
    @{N="LogDB";T="num"};@{N="LogTimer";T="num"};@{N="AlLoLo";T="num"};@{N="AlHiHi";T="num"}
    @{N="MbsTcpRegister";T="num"};@{N="MbsTcpCoef";T="num"};@{N="MbsTcpOffset";T="num"}
    @{N="EEN";T="num"};@{N="ETO";T="str"};@{N="ECC";T="str"};@{N="ESU";T="str"};@{N="EAT";T="str"}
    @{N="ESH";T="num"};@{N="SEN";T="num"};@{N="STO";T="str"};@{N="SSU";T="str"}
    @{N="TEN";T="num"};@{N="TSU";T="str"};@{N="FEN";T="num"};@{N="FFN";T="str"};@{N="FCO";T="str"}
    @{N="KPI";T="num"};@{N="UseCustomUnit";T="num"};@{N="Type";T="num"};@{N="Unit";T="str"}
    @{N="AlStat";T="num"};@{N="ChangeTime";T="str"};@{N="TagValue";T="num"}
    @{N="TagQuality";T="num"};@{N="AlType";T="num"}
)

# Float column indices (0-based) — formatted with 6 decimal places
$Script:EWON_FLOAT_COLS = @(6, 7, 21, 22, 24, 32, 37, 38)

# Unit lookup cache
$Script:UnitLookup = $null

function Get-UnitLookup {
    if ($Script:UnitLookup) { return $Script:UnitLookup }

    $Script:UnitLookup = @{}
    $jsonRaw = $null

    # Priority 1: Embedded JSON (release mode — single .ps1)
    if ($Script:EmbeddedUnitsJson) {
        $jsonRaw = $Script:EmbeddedUnitsJson
    } else {
        # Priority 2: External file (dev mode — modules folder)
        $jsonPath = Join-Path $PSScriptRoot "..\data\units.json"
        if (Test-Path $jsonPath) {
            $jsonRaw = Get-Content $jsonPath -Raw -Encoding UTF8
        }
    }

    if ($jsonRaw) {
        try {
            $units = $jsonRaw | ConvertFrom-Json
            foreach ($u in $units) {
                if ($u.displayName -and $u.uneceCode) {
                    $Script:UnitLookup[$u.displayName] = $u.uneceCode
                }
            }
        } catch {}
    }
    return $Script:UnitLookup
}

function Resolve-EwonUnit {
    param([string]$Unite)

    if (-not $Unite) { return @{ Unit = ""; UseCustomUnit = 0 } }

    $lookup = Get-UnitLookup
    $code = $lookup[$Unite]
    if ($code) {
        return @{ Unit = $code; UseCustomUnit = 0 }
    }
    return @{ Unit = $Unite; UseCustomUnit = 1 }
}

function Get-EwonS7Address {
    param(
        [int]$DbNumber,
        [string]$Offset,
        [string]$DataType,
        [string]$IpAddress,
        [string]$Tsap
    )

    if (-not $Offset) { return "" }

    # Parse offset "545.2" → byte=545, bit=2
    $parts = $Offset.Split('.')
    $byteOffset = [int]$parts[0]
    $bitNumber = if ($parts.Count -gt 1) { [int]$parts[1] } else { 0 }

    # Clean datatype (remove quotes from UDT names)
    $dt = $DataType.Trim('"')
    $suffix = $Script:S7_EWON_FORMAT[$dt]
    if (-not $suffix) { $suffix = "W" }

    $address = "DB${DbNumber}${suffix}${byteOffset}"

    # Bool: append #bit
    if ($dt -eq "Bool") {
        $address += "#${bitNumber}"
    }

    # Communication params: ,ISOTCP,{ip},{tsap}
    if ($IpAddress) {
        $address += ",ISOTCP,${IpAddress}"
    }
    if ($Tsap) {
        $address += ",${Tsap}"
    }

    return $address
}

function Format-EwonStr {
    param([string]$Value)
    if (-not $Value) { return '""' }
    $escaped = $Value.Replace('"', '""')
    return "`"$escaped`""
}

function Format-EwonFloat {
    param($Value)
    if ($null -eq $Value) { return "" }
    return ([double]$Value).ToString("F6")
}

function Write-VarLstHeader {
    param([System.IO.StreamWriter]$Writer)
    $header = ($Script:VAR_LST_COLUMNS | ForEach-Object { "`"$($_.N)`"" }) -join ";"
    $Writer.Write("$header`r`n")
}

function Write-VarLstMemberRow {
    param(
        [System.IO.StreamWriter]$Writer,
        [hashtable]$Member,
        [hashtable]$PlcInfo,
        [hashtable]$EwonConfig
    )

    $dt = $Member.DataType.Trim('"')
    $ewonType = $Script:EWON_TYPE_MAP[$dt]
    if ($null -eq $ewonType) { $ewonType = 0 }

    # Name: {repere}.{memberPath}
    $tagName = if ($EwonConfig.Repere) { "$($EwonConfig.Repere).$($Member.Name)" } else { $Member.Name }

    # Address
    $address = Get-EwonS7Address -DbNumber $Member.DbNumber -Offset $Member.Offset -DataType $dt -IpAddress $PlcInfo.IpAddress -Tsap $PlcInfo.Tsap

    # Unit resolution (Member may not have a Unit key)
    $memberUnit = if ($Member.ContainsKey('Unit')) { $Member.Unit } else { "" }
    $unitInfo = Resolve-EwonUnit -Unite $memberUnit

    # Build 62-column values array (matching VAR_LST_COLUMNS order)
    $vals = @(
        ""                                          #  0 Id (auto)
        (Format-EwonStr $tagName)                   #  1 Name
        (Format-EwonStr $Member.Comment)            #  2 Description
        (Format-EwonStr "S7300")                    #  3 ServerName
        (Format-EwonStr $EwonConfig.Topic)          #  4 TopicName
        (Format-EwonStr $address)                   #  5 Address
        (Format-EwonFloat 1)                        #  6 Coef
        (Format-EwonFloat 0)                        #  7 Offset (scaling)
        "1"                                         #  8 LogEnabled
        "0"                                         #  9 AlEnabled
        "0"                                         # 10 AlBool
        "0"                                         # 11 MemTag
        "0"                                         # 12 MbsTcpEnabled
        "0"                                         # 13 MbsTcpFloat
        "0"                                         # 14 SnmpEnabled
        "0"                                         # 15 RTLogEnabled
        "0"                                         # 16 AlAutoAck
        "0"                                         # 17 ForceRO
        "1"                                         # 18 SnmpOID
        "0"                                         # 19 AutoType
        '""'                                        # 20 AlHint
        (Format-EwonFloat 0)                        # 21 AlHigh
        (Format-EwonFloat 0)                        # 22 AlLow
        "0"                                         # 23 AlTimeDB
        (Format-EwonFloat 0)                        # 24 AlLevelDB
        "0"                                         # 25 IVGroupA
        "0"                                         # 26 IVGroupB
        "0"                                         # 27 IVGroupC
        "0"                                         # 28 IVGroupD
        "$($EwonConfig.PageId)"                     # 29 PageId
        "600"                                       # 30 RTLogWindow
        "10"                                        # 31 RTLogTimer
        (Format-EwonFloat (-1))                     # 32 LogDB
        "60"                                        # 33 LogTimer
        ""                                          # 34 AlLoLo
        ""                                          # 35 AlHiHi
        "1"                                         # 36 MbsTcpRegister
        (Format-EwonFloat 1)                        # 37 MbsTcpCoef
        (Format-EwonFloat 0)                        # 38 MbsTcpOffset
        ""                                          # 39 EEN
        '""'                                        # 40 ETO
        '""'                                        # 41 ECC
        '""'                                        # 42 ESU
        '""'                                        # 43 EAT
        ""                                          # 44 ESH
        ""                                          # 45 SEN
        '""'                                        # 46 STO
        '""'                                        # 47 SSU
        ""                                          # 48 TEN
        '""'                                        # 49 TSU
        ""                                          # 50 FEN
        '""'                                        # 51 FFN
        '""'                                        # 52 FCO
        "0"                                         # 53 KPI
        "$($unitInfo.UseCustomUnit)"                # 54 UseCustomUnit
        "$ewonType"                                 # 55 Type
        (Format-EwonStr $unitInfo.Unit)             # 56 Unit
        "0"                                         # 57 AlStat
        '""'                                        # 58 ChangeTime
        "0"                                         # 59 TagValue
        "65472"                                     # 60 TagQuality
        "0"                                         # 61 AlType
    )

    $Writer.Write(($vals -join ";") + "`r`n")
}

# =================== PCVUE ARCHITECT FORMAT ===================

function Get-PcVueOffset {
    param([string]$SiemensOffset, [string]$DataType)

    if (-not $SiemensOffset) { return @{ Decalage = 0; WBIT = 0 } }

    $parts = $SiemensOffset.Split('.')
    $byte = [int]$parts[0]
    $bit = if ($parts.Count -gt 1) { [int]$parts[1] } else { 0 }

    $dt = $DataType.Trim('"')
    if ($dt -eq "Bool") {
        # Word-swap: even byte → +1, odd byte → -1
        if ($byte % 2 -eq 0) { $pcvueByte = $byte + 1 } else { $pcvueByte = $byte - 1 }
        return @{ Decalage = $pcvueByte; WBIT = $bit }
    }

    return @{ Decalage = $byte; WBIT = 0 }
}

function Write-PcVueCsvHeader {
    param([System.IO.StreamWriter]$Writer)
    $Writer.WriteLine("Nom;Adresse;Type;Description;Decalage;WBIT;Trame;Colonne1;Colonne2;Colonne3;Colonne4;Colonne5;Colonne6;Colonne7;Colonne8")
}

function Write-PcVueMemberRow {
    param(
        [System.IO.StreamWriter]$Writer,
        [hashtable]$Member,
        [bool]$IsOptimized
    )

    $dt = $Member.DataType.Trim('"')
    $comment = ($Member.Comment -replace '"', '""')
    if ($comment -match '[;"]') { $comment = "`"$comment`"" }

    $name = ($Member.Name -replace '"', '""')
    if ($name -match '[;"]') { $name = "`"$name`"" }

    if ($IsOptimized) {
        $decalage = ""
        $wbit = ""
    } else {
        $pcvue = Get-PcVueOffset -SiemensOffset $Member.Offset -DataType $dt
        $decalage = $pcvue.Decalage
        $wbit = $pcvue.WBIT
    }

    $trame = "DB$($Member.DbNumber)"

    $Writer.WriteLine("$name;;$dt;$comment;$decalage;$wbit;$trame;;;;;;;;")
}

# =================== MAIN EXPORT ORCHESTRATION ===================

function Invoke-TableExport {
    param(
        [array]$SelectedBlocks,
        [string]$OutputFolder,
        [string]$Format = "CSV",
        [hashtable]$EwonConfig = $null,
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

    $result.Format = $Format

    $state = Get-AppState
    $plcInfoList = $state.PlcDeviceInfoList

    # Load unit lookup for Ewon format
    if ($Format -eq "EWON") {
        Get-UnitLookup | Out-Null
    }

    # Create temp folder for XML exports (in system temp)
    $tempFolder = Join-Path ([System.IO.Path]::GetTempPath()) "TIA_Export_$([guid]::NewGuid().ToString('N'))"
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
        $typeCommentCache = @{}  # Cache: TypeName → @{ "member.path" → "comment" }

        foreach ($plcIdx in ($blocksByPlc.Keys | Sort-Object)) {
            $plcBlocks = $blocksByPlc[$plcIdx]

            # Get PLC info
            $plcInfo = $plcInfoList | Where-Object { $_.PlcIndex -eq $plcIdx } | Select-Object -First 1
            if (-not $plcInfo) {
                $plcInfo = @{ PlcIndex = $plcIdx; Name = "PLC_$plcIdx"; IpAddress = ""; Rack = 0; Slot = 2; Tsap = "03.02" }
            }

            $safeName = $plcInfo.Name -replace '[\\/:*?"<>|]', '_'
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

            # PCVUE: create timestamped subfolder (1 file per DB)
            # CSV/EWON: single file per PLC
            $writer = $null
            $pcvueFolder = $null

            if ($Format -eq "PCVUE") {
                $pcvueFolder = Join-Path $OutputFolder "PcVue_${safeName}_${timestamp}"
                New-Item -ItemType Directory -Path $pcvueFolder -Force | Out-Null
                $result.OutputFolder = $pcvueFolder
            } elseif ($Format -eq "EWON") {
                $csvFileName = "var_lst_${safeName}.csv"
                $csvPath = Join-Path $OutputFolder $csvFileName
                $encoding = [System.Text.Encoding]::GetEncoding("iso-8859-1")
                $writer = New-Object System.IO.StreamWriter($csvPath, $false, $encoding)
                Write-VarLstHeader -Writer $writer
            } else {
                $csvFileName = "${safeName}_Export_${timestamp}.csv"
                $csvPath = Join-Path $OutputFolder $csvFileName
                $encoding = New-Object System.Text.UTF8Encoding($true)
                $writer = New-Object System.IO.StreamWriter($csvPath, $false, $encoding)
                Write-CsvHeader -Writer $writer -PlcInfo $plcInfo
            }

            try {
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
                        try {
                            $xmlPath = Export-BlockToXml -Block $block -TempFolder $tempFolder
                        } catch {
                            if ($_.Exception.Message -match "Inconsistent") {
                                $result.Errors += "$($dbInfo.Name): $(T 'MsgInconsistentBlock')"
                            } else {
                                $result.Errors += "$($dbInfo.Name): $($_.Exception.Message)"
                            }
                            $result.ErrorCount++
                            continue
                        }

                        # Parse SimaticML XML
                        $parseResult = Parse-SimaticMlMembers -XmlPath $xmlPath -DbNumber $dbInfo.Number

                        if ($parseResult.IsOptimized) {
                            $result.OptimizedDBs += "DB$($dbInfo.Number)_$($dbInfo.Name)"
                            $result.SuccessCount++
                            if ($OnProgress) {
                                & $OnProgress $currentBlock $totalBlocks $dbInfo.Name
                            }
                            continue
                        }

                        # Resolve comments from source types (FB + UDTs)
                        [xml]$dbXml = Get-Content $xmlPath -Encoding UTF8
                        $instanceOfName = Get-InstanceOfName -XmlDoc $dbXml

                        # Collect all type names that need comment resolution
                        $typeNames = @{}
                        foreach ($m in $parseResult.Members) {
                            if ([string]::IsNullOrEmpty($m.Comment)) {
                                $src = if ($m.UdtType) { $m.UdtType } elseif ($instanceOfName) { $instanceOfName } else { $null }
                                if ($src -and -not $typeNames.ContainsKey($src)) { $typeNames[$src] = @{} }
                            }
                        }

                        # Export each type and extract comments (cached per type)
                        foreach ($typeName in @($typeNames.Keys)) {
                            if (-not $typeCommentCache.ContainsKey($typeName)) {
                                try {
                                    $typeObj = $null
                                    foreach ($plcSw in $state.PlcSoftwareList) {
                                        $typeObj = Find-BlockByName -BlockGroup $plcSw.BlockGroup -Name $typeName
                                        if ($typeObj) { break }
                                        try {
                                            $typeObj = Find-TypeByName -TypeGroup $plcSw.TypeGroup -Name $typeName
                                        } catch {}
                                        if ($typeObj) { break }
                                    }
                                    if ($typeObj) {
                                        $typeXmlPath = Export-BlockToXml -Block $typeObj -TempFolder $tempFolder
                                        [xml]$typeXml = Get-Content $typeXmlPath -Encoding UTF8
                                        $typeSections = $typeXml.GetElementsByTagName("Section")
                                        $tc = @{}
                                        foreach ($sec in $typeSections) {
                                            $sc = Get-CommentsFromXml -ParentNode $sec
                                            foreach ($k in $sc.Keys) { $tc[$k] = $sc[$k] }
                                        }
                                        $typeCommentCache[$typeName] = $tc
                                    } else {
                                        $typeCommentCache[$typeName] = @{}
                                    }
                                } catch {
                                    $typeCommentCache[$typeName] = @{}
                                }
                            }
                        }

                        # Inject resolved comments
                        foreach ($m in $parseResult.Members) {
                            if ([string]::IsNullOrEmpty($m.Comment)) {
                                $src = if ($m.UdtType) { $m.UdtType } elseif ($instanceOfName) { $instanceOfName } else { $null }
                                if ($src -and $typeCommentCache.ContainsKey($src) -and $typeCommentCache[$src].ContainsKey($m.UdtPath)) {
                                    $m.Comment = $typeCommentCache[$src][$m.UdtPath]
                                }
                            }
                        }

                        # PCVUE: open a new writer per DB
                        $pcvueWriter = $null
                        if ($Format -eq "PCVUE") {
                            $safeDbName = $dbInfo.Name -replace '[\\/:*?"<>|]', '_'
                            $pcvueFile = Join-Path $pcvueFolder "DB$($dbInfo.Number)_${safeDbName}.csv"
                            $pcvueEncoding = New-Object System.Text.UTF8Encoding($true)
                            $pcvueWriter = New-Object System.IO.StreamWriter($pcvueFile, $false, $pcvueEncoding)
                            Write-PcVueCsvHeader -Writer $pcvueWriter
                        }

                        try {
                            # Write member rows
                            foreach ($member in $parseResult.Members) {
                                try {
                                    if ($Format -eq "PCVUE") {
                                        Write-PcVueMemberRow -Writer $pcvueWriter -Member $member -IsOptimized $false
                                    } elseif ($Format -eq "EWON") {
                                        Write-VarLstMemberRow -Writer $writer -Member $member -PlcInfo $plcInfo -EwonConfig $EwonConfig
                                    } else {
                                        Write-CsvMemberRow -Writer $writer -Member $member -IsOptimized $false
                                    }
                                } catch {
                                    $memberName = if ($member -and $member.ContainsKey('Name')) { $member.Name } else { "?" }
                                    $result.Errors += "DB$($dbInfo.Number).$memberName : $($_.Exception.Message) [$($_.InvocationInfo.ScriptLineNumber)]"
                                }
                            }
                        } finally {
                            if ($pcvueWriter) {
                                $pcvueWriter.Flush()
                                $pcvueWriter.Close()
                                $result.Files += $pcvueFile
                            }
                        }

                        if ($parseResult.Members.Count -eq 0) {
                            $result.Errors += "DB$($dbInfo.Number)_$($dbInfo.Name): 0 members parsed"
                        }

                        $result.SuccessCount++

                        if ($OnProgress) {
                            & $OnProgress $currentBlock $totalBlocks $dbInfo.Name
                        }

                    } catch {
                        if ($_.Exception.Message -match "Inconsistent") {
                            $result.Errors += "$($dbInfo.Name): $(T 'MsgInconsistentBlock')"
                        } else {
                            $result.Errors += "$($dbInfo.Name): $($_.Exception.Message) [line $($_.InvocationInfo.ScriptLineNumber)]"
                        }
                        $result.ErrorCount++
                    }
                }

                if ($writer) {
                    $writer.Flush()
                    $result.Files += $csvPath
                }
            } finally {
                if ($writer) { $writer.Close() }
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

    $doneKey = switch ($Result.Format) {
        "EWON"  { "MsgExportEwonDone" }
        "PCVUE" { "MsgExportPcVueDone" }
        default { "MsgExportCsvDone" }
    }
    $msg = (T $doneKey) + "`n`n"
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
