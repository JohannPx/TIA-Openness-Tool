# UIHelpers.ps1 - Reusable WPF helper functions
# Pattern from ewon-flexy-config\scripts\modules\UIHelpers.ps1

Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

function Set-StatusBanner {
    param(
        [System.Windows.Controls.Border]$Banner,
        [System.Windows.Controls.TextBlock]$TextBlock,
        [string]$Text,
        [string]$Type  # "success" | "error" | "info" | "warning"
    )

    $brush = [System.Windows.Media.BrushConverter]::new()
    $TextBlock.Text = $Text

    switch ($Type) {
        "success" {
            $Banner.Background = $brush.ConvertFrom("#E8F8E8")
            $TextBlock.Foreground = $brush.ConvertFrom("#27AE60")
        }
        "error" {
            $Banner.Background = $brush.ConvertFrom("#FDE8E8")
            $TextBlock.Foreground = $brush.ConvertFrom("#E74C3C")
        }
        "warning" {
            $Banner.Background = $brush.ConvertFrom("#FEF3C7")
            $TextBlock.Foreground = $brush.ConvertFrom("#F59E0B")
        }
        default {  # info
            $Banner.Background = $brush.ConvertFrom("#E8F0FE")
            $TextBlock.Foreground = $brush.ConvertFrom("#1A5276")
        }
    }

    $Banner.Visibility = [System.Windows.Visibility]::Visible
}

function Set-ElementVisibility {
    param(
        [System.Windows.UIElement]$Element,
        [bool]$Visible
    )
    $Element.Visibility = if ($Visible) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
}

function New-SectionHeader {
    param([string]$Title)

    $border = New-Object System.Windows.Controls.Border
    $border.Margin = [System.Windows.Thickness]::new(0, 12, 0, 4)
    $border.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#1A5276")
    $border.BorderThickness = [System.Windows.Thickness]::new(0, 0, 0, 1)
    $border.Padding = [System.Windows.Thickness]::new(0, 0, 0, 4)

    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Title
    $tb.FontSize = 13
    $tb.FontWeight = [System.Windows.FontWeights]::SemiBold
    $tb.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#1A5276")

    $border.Child = $tb
    return $border
}

function New-DataBlockListItem {
    param([hashtable]$Block)

    $brush = [System.Windows.Media.BrushConverter]::new()

    # Main container
    $border = New-Object System.Windows.Controls.Border
    $border.Padding = [System.Windows.Thickness]::new(8, 6, 8, 6)
    $border.Margin = [System.Windows.Thickness]::new(0, 1, 0, 1)
    $border.Background = [System.Windows.Media.Brushes]::White
    $border.BorderBrush = $brush.ConvertFrom("#E2E8F0")
    $border.BorderThickness = [System.Windows.Thickness]::new(0, 0, 0, 1)

    $grid = New-Object System.Windows.Controls.Grid

    # Columns: Checkbox(30) | Index(40) | Name(*) | Number(80) | TypeBadge(80) | PLC(60)
    @(30, 40, -1, 80, 80, 60) | ForEach-Object {
        $col = New-Object System.Windows.Controls.ColumnDefinition
        if ($_ -eq -1) {
            $col.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        } else {
            $col.Width = [System.Windows.GridLength]::new($_)
        }
        $grid.ColumnDefinitions.Add($col)
    }

    # Checkbox
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.IsChecked = $Block.IsSelected
    $cb.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $cb.Tag = $Block
    $cb.Add_Checked({ param($s,$e); $s.Tag.IsSelected = $true })
    $cb.Add_Unchecked({ param($s,$e); $s.Tag.IsSelected = $false })
    [System.Windows.Controls.Grid]::SetColumn($cb, 0)
    $grid.Children.Add($cb) | Out-Null

    # Index
    $txtIdx = New-Object System.Windows.Controls.TextBlock
    $txtIdx.Text = $Block.Index.ToString()
    $txtIdx.FontSize = 11
    $txtIdx.Foreground = $brush.ConvertFrom("#999")
    $txtIdx.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    [System.Windows.Controls.Grid]::SetColumn($txtIdx, 1)
    $grid.Children.Add($txtIdx) | Out-Null

    # Name
    $txtName = New-Object System.Windows.Controls.TextBlock
    $txtName.Text = $Block.Name
    $txtName.FontSize = 12
    $txtName.FontWeight = [System.Windows.FontWeights]::Medium
    $txtName.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $txtName.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis
    [System.Windows.Controls.Grid]::SetColumn($txtName, 2)
    $grid.Children.Add($txtName) | Out-Null

    # Number
    $txtNum = New-Object System.Windows.Controls.TextBlock
    $txtNum.Text = "DB$($Block.Number)"
    $txtNum.FontSize = 11
    $txtNum.Foreground = $brush.ConvertFrom("#666")
    $txtNum.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    [System.Windows.Controls.Grid]::SetColumn($txtNum, 3)
    $grid.Children.Add($txtNum) | Out-Null

    # Type badge
    $badge = New-Object System.Windows.Controls.Border
    $badge.Background = $brush.ConvertFrom($Block.BadgeBackgroundColor)
    $badge.CornerRadius = [System.Windows.CornerRadius]::new(3)
    $badge.Padding = [System.Windows.Thickness]::new(6, 2, 6, 2)
    $badge.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $badge.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

    $txtBadge = New-Object System.Windows.Controls.TextBlock
    $txtBadge.Text = $Block.DisplayType
    $txtBadge.FontSize = 10
    $txtBadge.FontWeight = [System.Windows.FontWeights]::SemiBold
    $txtBadge.Foreground = $brush.ConvertFrom($Block.BadgeColor)
    $badge.Child = $txtBadge
    [System.Windows.Controls.Grid]::SetColumn($badge, 4)
    $grid.Children.Add($badge) | Out-Null

    # PLC index
    $txtPlc = New-Object System.Windows.Controls.TextBlock
    $txtPlc.Text = (T "LblPlc") -f $Block.PlcIndex
    $txtPlc.FontSize = 10
    $txtPlc.Foreground = $brush.ConvertFrom("#999")
    $txtPlc.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $txtPlc.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    [System.Windows.Controls.Grid]::SetColumn($txtPlc, 5)
    $grid.Children.Add($txtPlc) | Out-Null

    $border.Child = $grid
    return $border
}

function New-InstanceListItem {
    param([hashtable]$Instance)

    $brush = [System.Windows.Media.BrushConverter]::new()

    $border = New-Object System.Windows.Controls.Border
    $border.Padding = [System.Windows.Thickness]::new(12, 8, 12, 8)
    $border.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)
    $border.Background = [System.Windows.Media.Brushes]::White
    $border.BorderBrush = $brush.ConvertFrom("#E2E8F0")
    $border.BorderThickness = [System.Windows.Thickness]::new(1)
    $border.CornerRadius = [System.Windows.CornerRadius]::new(4)

    $sp = New-Object System.Windows.Controls.StackPanel

    $txtMain = New-Object System.Windows.Controls.TextBlock
    $txtMain.Text = $Instance.DisplayText
    $txtMain.FontSize = 13
    $txtMain.FontWeight = [System.Windows.FontWeights]::Medium
    $sp.Children.Add($txtMain) | Out-Null

    if ($Instance.ProjectName) {
        $txtProj = New-Object System.Windows.Controls.TextBlock
        $txtProj.Text = (T "LblProject") -f $Instance.ProjectName
        $txtProj.FontSize = 11
        $txtProj.Foreground = $brush.ConvertFrom("#666")
        $txtProj.Margin = [System.Windows.Thickness]::new(0, 2, 0, 0)
        $sp.Children.Add($txtProj) | Out-Null
    }

    $border.Child = $sp
    $border.Tag = $Instance
    return $border
}
