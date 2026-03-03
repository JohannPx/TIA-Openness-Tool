# UI.ps1 - XAML definition, window initialization, event wiring
# Pattern from ewon-flexy-config\scripts\modules\UI.ps1

$Script:MainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="TIA Openness Tool"
        Width="1100" Height="720"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize"
        MinWidth="900" MinHeight="600"
        UseLayoutRounding="True"
        SnapsToDevicePixels="True"
        Background="#FAFAFA">
  <Window.Resources>
    <Style x:Key="PageHeader" TargetType="TextBlock">
      <Setter Property="FontSize" Value="17"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#1A5276"/>
      <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>
    <Style x:Key="SubText" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#666"/>
      <Setter Property="FontSize" Value="11"/>
    </Style>
    <Style x:Key="LangBtn" TargetType="Button">
      <Setter Property="Width" Value="52"/>
      <Setter Property="Height" Value="30"/>
      <Setter Property="Margin" Value="3,0"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="BorderThickness" Value="2"/>
      <Setter Property="BorderBrush" Value="Transparent"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="VerticalContentAlignment" Value="Stretch"/>
    </Style>
  </Window.Resources>

  <DockPanel>
    <!-- =================== TOP BAR =================== -->
    <Border DockPanel.Dock="Top" Background="#1A5276" Padding="18,10">
      <DockPanel>
        <!-- Left: App title -->
        <StackPanel DockPanel.Dock="Left" Orientation="Horizontal" VerticalAlignment="Center">
          <Border Width="34" Height="34" Background="White" CornerRadius="4" Margin="0,0,12,0">
            <TextBlock Text="DB" FontSize="14" FontWeight="Bold" Foreground="#1A5276"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel VerticalAlignment="Center">
            <TextBlock x:Name="txtAppTitle" Text="TIA Openness Tool" FontSize="16"
                       FontWeight="Bold" Foreground="White"/>
            <TextBlock x:Name="txtAppSubtitle" Text="Export DataBlocks TIA Portal" FontSize="10"
                       Foreground="#A0C4E0"/>
          </StackPanel>
        </StackPanel>

        <!-- Right: Connection status -->
        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <Border x:Name="brdConnected" Visibility="Collapsed" Background="#27AE60"
                  CornerRadius="3" Padding="14,5">
            <StackPanel Orientation="Horizontal">
              <Ellipse Width="8" Height="8" Fill="White" Margin="0,0,8,0"/>
              <TextBlock x:Name="txtConnectedLabel" Text="Connecte"
                         Foreground="White" FontSize="11" FontWeight="SemiBold"/>
            </StackPanel>
          </Border>
          <Border x:Name="brdDisconnected" Visibility="Visible" Background="#6B7280"
                  CornerRadius="3" Padding="14,5">
            <StackPanel Orientation="Horizontal">
              <Ellipse Width="8" Height="8" Fill="White" Margin="0,0,8,0"/>
              <TextBlock x:Name="txtDisconnectedLabel" Text="Deconnecte"
                         Foreground="White" FontSize="11" FontWeight="SemiBold"/>
            </StackPanel>
          </Border>
        </StackPanel>

        <!-- Center: Language flags -->
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
          <TextBlock x:Name="txtLangLabel" Text="Langue :" VerticalAlignment="Center"
                     Margin="0,0,8,0" FontSize="11" Foreground="#A0C4E0"/>
          <Button x:Name="btnLangFR" Style="{StaticResource LangBtn}" Tag="FR" ToolTip="Francais">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Rectangle Grid.Column="0" Fill="#002395"/>
              <Rectangle Grid.Column="1" Fill="White"/>
              <Rectangle Grid.Column="2" Fill="#ED2939"/>
            </Grid>
          </Button>
          <Button x:Name="btnLangEN" Style="{StaticResource LangBtn}" Tag="EN" ToolTip="English">
            <Grid Background="#012169">
              <Rectangle Fill="White" Width="10" HorizontalAlignment="Center"/>
              <Rectangle Fill="White" Height="8" VerticalAlignment="Center"/>
              <Rectangle Fill="#CF142B" Width="5" HorizontalAlignment="Center"/>
              <Rectangle Fill="#CF142B" Height="4" VerticalAlignment="Center"/>
            </Grid>
          </Button>
          <Button x:Name="btnLangES" Style="{StaticResource LangBtn}" Tag="ES" ToolTip="Espanol">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="*"/><RowDefinition Height="2*"/><RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Rectangle Grid.Row="0" Fill="#AA151B"/>
              <Rectangle Grid.Row="1" Fill="#F1BF00"/>
              <Rectangle Grid.Row="2" Fill="#AA151B"/>
            </Grid>
          </Button>
          <Button x:Name="btnLangIT" Style="{StaticResource LangBtn}" Tag="IT" ToolTip="Italiano">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Rectangle Grid.Column="0" Fill="#009246"/>
              <Rectangle Grid.Column="1" Fill="White"/>
              <Rectangle Grid.Column="2" Fill="#CE2B37"/>
            </Grid>
          </Button>
        </StackPanel>
      </DockPanel>
    </Border>

    <!-- =================== LEFT SIDEBAR =================== -->
    <Border DockPanel.Dock="Left" Width="200" Background="#F8FAFC" BorderBrush="#E2E8F0"
            BorderThickness="0,0,1,0">
      <DockPanel>
        <!-- Version selector -->
        <StackPanel DockPanel.Dock="Top" Margin="12,16,12,12">
          <TextBlock x:Name="txtVersionLabel" Text="Version TIA Portal :" FontSize="11"
                     Foreground="#666" Margin="0,0,0,4"/>
          <ComboBox x:Name="cbTiaVersion" Height="30" FontSize="12"/>
          <TextBlock x:Name="txtVersionInfo" Text="" FontSize="9" Foreground="#999"
                     TextWrapping="Wrap" Margin="0,4,0,0"/>
        </StackPanel>

        <Border DockPanel.Dock="Top" Height="1" Background="#E2E8F0" Margin="12,0"/>

        <!-- Navigation buttons -->
        <StackPanel DockPanel.Dock="Top" Margin="0,8,0,0">
          <Button x:Name="btnNavConnection" Height="44" Background="#EAF2F8"
                  BorderThickness="0" HorizontalContentAlignment="Left" Padding="16,0" Cursor="Hand">
            <TextBlock x:Name="txtNavConnection" Text="Connexion" FontSize="13"
                       Foreground="#1A5276" FontWeight="SemiBold"/>
          </Button>
          <Button x:Name="btnNavExport" Height="44" Background="Transparent"
                  BorderThickness="0" HorizontalContentAlignment="Left" Padding="16,0" Cursor="Hand">
            <TextBlock x:Name="txtNavExport" Text="Export DataBlocks" FontSize="13"
                       Foreground="#4A5568"/>
          </Button>
        </StackPanel>

        <Control/>
      </DockPanel>
    </Border>

    <!-- =================== MAIN CONTENT =================== -->
    <Grid Margin="20">

      <!-- ===== PAGE: Connection ===== -->
      <Grid x:Name="pnlConnection" Visibility="Visible">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" x:Name="txtConnTitle" Text="Connexion TIA Portal"
                   Style="{StaticResource PageHeader}"/>

        <!-- Scan button -->
        <Button Grid.Row="1" x:Name="btnScan" Content="Scanner les instances TIA Portal"
                Height="40" FontSize="13" Cursor="Hand" Margin="0,0,0,16"
                Background="White" BorderBrush="#CBD5E0" BorderThickness="1"/>

        <!-- Instance list -->
        <Border Grid.Row="2" Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
                CornerRadius="4" Padding="16">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <DockPanel Grid.Row="0" Margin="0,0,0,8">
              <TextBlock x:Name="txtInstancesLabel" Text="Instances disponibles"
                         FontWeight="SemiBold" FontSize="13" Foreground="#4A5568"
                         DockPanel.Dock="Left" VerticalAlignment="Center"/>
              <Border x:Name="brdScanStatus" DockPanel.Dock="Right" Padding="10,4"
                      CornerRadius="3" Background="#E8F0FE" Visibility="Collapsed">
                <TextBlock x:Name="txtScanStatus" FontSize="11" Foreground="#1A5276"/>
              </Border>
              <Control/>
            </DockPanel>
            <ListBox Grid.Row="1" x:Name="lbInstances" BorderThickness="0"
                     Background="Transparent" HorizontalContentAlignment="Stretch"/>
          </Grid>
        </Border>

        <!-- Connect/Disconnect buttons -->
        <Border Grid.Row="3" Margin="0,12,0,0">
          <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
            <Button x:Name="btnConnect" Content="Se connecter" Width="180" Height="40"
                    FontSize="13" FontWeight="SemiBold" Cursor="Hand"
                    Background="#1A5276" Foreground="White" BorderThickness="0" Margin="0,0,8,0"/>
            <Button x:Name="btnDisconnect" Content="Se deconnecter" Width="180" Height="40"
                    FontSize="13" Cursor="Hand" IsEnabled="False"
                    Background="White" BorderBrush="#CBD5E0" BorderThickness="1"/>
          </StackPanel>
        </Border>
      </Grid>

      <!-- ===== PAGE: Export DataBlocks ===== -->
      <Grid x:Name="pnlExport" Visibility="Collapsed">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" x:Name="txtExportTitle" Text="Export DataBlocks"
                   Style="{StaticResource PageHeader}"/>

        <!-- Load button -->
        <Button Grid.Row="1" x:Name="btnLoadDBs" Content="Charger les DataBlocks"
                Height="40" FontSize="13" Cursor="Hand" Margin="0,0,0,12"
                Background="White" BorderBrush="#CBD5E0" BorderThickness="1"/>

        <!-- Toolbar -->
        <Grid Grid.Row="2" Margin="0,0,0,8">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Button x:Name="btnSelectAll" Grid.Column="0" Content="Tout selectionner"
                  Height="30" FontSize="11" Cursor="Hand" Padding="12,0" Margin="0,0,6,0"
                  Background="White" BorderBrush="#CBD5E0" BorderThickness="1"/>
          <Button x:Name="btnDeselectAll" Grid.Column="1" Content="Tout deselectionner"
                  Height="30" FontSize="11" Cursor="Hand" Padding="12,0" Margin="0,0,6,0"
                  Background="White" BorderBrush="#CBD5E0" BorderThickness="1"/>
          <CheckBox x:Name="chkHideInstance" Grid.Column="2" Content="Masquer les DBs d'instance"
                    VerticalAlignment="Center" FontSize="12" IsChecked="True" Margin="12,0"/>
          <Border Grid.Column="3" Background="#EDF2F7" Padding="12,4" CornerRadius="3">
            <TextBlock x:Name="txtDBCount" Text="0 DB(s)" FontWeight="SemiBold"
                       Foreground="#4A5568" FontSize="12"/>
          </Border>
        </Grid>

        <!-- PLC Info Panel (visible after loading DataBlocks) -->
        <Border x:Name="brdPlcInfo" Grid.Row="3" Visibility="Collapsed"
                Background="#F0F9FF" BorderBrush="#BFDBFE" BorderThickness="1"
                CornerRadius="4" Padding="10" Margin="0,0,0,8">
          <StackPanel x:Name="spPlcInfo"/>
        </Border>

        <!-- DataBlock list -->
        <Border Grid.Row="4" Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
                CornerRadius="4">
          <ListBox x:Name="lbDataBlocks" BorderThickness="0" Background="Transparent"
                   HorizontalContentAlignment="Stretch"
                   VirtualizingStackPanel.IsVirtualizing="True"
                   ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
        </Border>

        <!-- Export section -->
        <Border Grid.Row="5" Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
                CornerRadius="4" Padding="16" Margin="0,12,0,0">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
              <TextBlock x:Name="txtExportFolderLabel" Text="Dossier :" FontSize="12"
                         VerticalAlignment="Center" Margin="0,0,8,0"/>
              <TextBlock x:Name="txtExportFolder" Text="Bureau (par defaut)" FontSize="12"
                         Foreground="#666" VerticalAlignment="Center"
                         TextTrimming="CharacterEllipsis" MaxWidth="400"/>
              <Button x:Name="btnBrowseFolder" Content="..." Width="32" Height="28"
                      Margin="8,0,0,0" Cursor="Hand" FontWeight="Bold"
                      Background="White" BorderBrush="#CBD5E0" BorderThickness="1"/>
            </StackPanel>
            <Button Grid.Column="1" x:Name="btnExportCsv" Content="Exporter Table CSV"
                    Width="240" Height="40" FontSize="13" FontWeight="SemiBold" Cursor="Hand"
                    Background="#27AE60" Foreground="White" BorderThickness="0"/>
          </Grid>
        </Border>
      </Grid>

    </Grid>
  </DockPanel>
</Window>
'@

# =================== INITIALIZE MAIN WINDOW ===================

function Initialize-MainWindow {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    # Set unique AppUserModelID for taskbar
    try {
        Add-Type -Name Shell32AppId -Namespace Native -ErrorAction SilentlyContinue -MemberDefinition @'
            [DllImport("shell32.dll", SetLastError = true)]
            public static extern void SetCurrentProcessExplicitAppUserModelID(
                [MarshalAs(UnmanagedType.LPWStr)] string AppID);
'@
        [Native.Shell32AppId]::SetCurrentProcessExplicitAppUserModelID("TIA.Openness.Tool.1")
    } catch {}

    # Parse XAML
    [xml]$xaml = $Script:MainXaml
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $Script:ui_Window = [System.Windows.Markup.XamlReader]::Load($reader)

    # Bind all named elements to script variables with ui_ prefix
    $elementNames = @(
        "txtAppTitle", "txtAppSubtitle",
        "txtLangLabel", "btnLangFR", "btnLangEN", "btnLangES", "btnLangIT",
        "brdConnected", "brdDisconnected", "txtConnectedLabel", "txtDisconnectedLabel",
        "txtVersionLabel", "cbTiaVersion", "txtVersionInfo",
        "btnNavConnection", "txtNavConnection", "btnNavExport", "txtNavExport",
        "pnlConnection", "txtConnTitle",
        "btnScan", "txtInstancesLabel", "brdScanStatus", "txtScanStatus", "lbInstances",
        "btnConnect", "btnDisconnect",
        "pnlExport", "txtExportTitle",
        "btnLoadDBs", "btnSelectAll", "btnDeselectAll", "chkHideInstance", "txtDBCount",
        "lbDataBlocks",
        "brdPlcInfo", "spPlcInfo",
        "txtExportFolderLabel", "txtExportFolder", "btnBrowseFolder", "btnExportCsv"
    )
    foreach ($name in $elementNames) {
        $el = $Script:ui_Window.FindName($name)
        if ($el) {
            Set-Variable -Name "ui_$name" -Value $el -Scope Script
        }
    }

    # Detect and populate TIA versions
    Initialize-VersionSelector

    # Apply language
    Update-AllTexts

    # Wire events
    Register-NavigationEvents
    Register-LanguageEvents
    Register-ConnectionEvents
    Register-ExportEvents

    # Window close guard
    $Script:ui_Window.Add_Closing({
        param($sender, $e)
        if ((Get-AppState).IsExporting) {
            $result = [System.Windows.MessageBox]::Show(
                (T "MsgConfirmClose"), (T "MsgConfirm"), "YesNo", "Warning")
            if ($result -eq "No") { $e.Cancel = $true; return }
        }
        # Clean up TIA connection
        if ((Get-AppState).IsConnected) {
            Disconnect-TiaInstance
        }
    })

    return $Script:ui_Window
}

# =================== VERSION SELECTOR ===================

function Initialize-VersionSelector {
    $versions = @(Get-InstalledTiaVersions)
    Set-AppStateValue -Key "InstalledVersions" -Value $versions

    $Script:ui_cbTiaVersion.Items.Clear()
    foreach ($v in $versions) {
        $Script:ui_cbTiaVersion.Items.Add($v.Version) | Out-Null
    }

    if ($versions.Count -eq 0) {
        $Script:ui_txtVersionInfo.Text = T "LblNoVersion"
    } elseif ($versions.Count -eq 1) {
        $Script:ui_cbTiaVersion.SelectedIndex = 0
    } else {
        # Select latest version by default
        $Script:ui_cbTiaVersion.SelectedIndex = $versions.Count - 1
    }

    $Script:ui_cbTiaVersion.Add_SelectionChanged({
        $selectedVersion = $Script:ui_cbTiaVersion.SelectedItem
        if (-not $selectedVersion) { return }
        $versions = (Get-AppState).InstalledVersions
        $match = $versions | Where-Object { $_.Version -eq $selectedVersion }
        if ($match) {
            $loadResult = Initialize-TiaOpenness -DllPath $match.DllPath
            if ($loadResult.Success) {
                Set-AppStateValue -Key "SelectedVersion" -Value $selectedVersion
                $Script:ui_txtVersionInfo.Text = (T "LblDllLoaded") -f [System.IO.Path]::GetFileName($match.DllPath)
            } else {
                $Script:ui_txtVersionInfo.Text = $loadResult.Message
            }
        }
    })

    # Trigger initial load
    if ($Script:ui_cbTiaVersion.SelectedItem) {
        $match = $versions | Where-Object { $_.Version -eq $Script:ui_cbTiaVersion.SelectedItem }
        if ($match) {
            $loadResult = Initialize-TiaOpenness -DllPath $match.DllPath
            if ($loadResult.Success) {
                Set-AppStateValue -Key "SelectedVersion" -Value $match.Version
                $Script:ui_txtVersionInfo.Text = (T "LblDllLoaded") -f [System.IO.Path]::GetFileName($match.DllPath)
            }
        }
    }
}

# =================== LANGUAGE EVENTS ===================

function Register-LanguageEvents {
    foreach ($langCode in @("FR","EN","ES","IT")) {
        $btn = Get-Variable -Name "ui_btnLang$langCode" -Scope Script -ValueOnly
        $btn.Add_Click({
            param($sender, $e)
            $lang = $sender.Tag
            Set-Language $lang
            Update-AllTexts
            Update-LanguageButtonHighlight
        })
    }
    Update-LanguageButtonHighlight
}

function Update-LanguageButtonHighlight {
    $currentLang = Get-Language
    $brush = [System.Windows.Media.BrushConverter]::new()
    foreach ($langCode in @("FR","EN","ES","IT")) {
        $btn = Get-Variable -Name "ui_btnLang$langCode" -Scope Script -ValueOnly
        if ($langCode -eq $currentLang) {
            $btn.BorderBrush = $brush.ConvertFrom("#FFD700")
            $btn.BorderThickness = [System.Windows.Thickness]::new(2)
        } else {
            $btn.BorderBrush = [System.Windows.Media.Brushes]::Transparent
            $btn.BorderThickness = [System.Windows.Thickness]::new(2)
        }
    }
}

function Update-AllTexts {
    # App title bar
    $Script:ui_txtAppTitle.Text = T "AppTitle"
    $Script:ui_txtAppSubtitle.Text = T "AppSubtitle"
    $Script:ui_txtLangLabel.Text = T "LangLabel"

    # Connection status
    $Script:ui_txtConnectedLabel.Text = T "LblConnected"
    $Script:ui_txtDisconnectedLabel.Text = T "LblDisconnected"

    # Sidebar
    $Script:ui_txtVersionLabel.Text = T "LblVersion"
    $Script:ui_txtNavConnection.Text = T "NavConnection"
    $Script:ui_txtNavExport.Text = T "NavExport"

    # Connection page
    $Script:ui_txtConnTitle.Text = T "PageConnection"
    $Script:ui_btnScan.Content = T "BtnScan"
    $Script:ui_txtInstancesLabel.Text = T "LblInstances"
    $Script:ui_btnConnect.Content = T "BtnConnect"
    $Script:ui_btnDisconnect.Content = T "BtnDisconnect"

    # Export page
    $Script:ui_txtExportTitle.Text = T "PageExport"
    $Script:ui_btnLoadDBs.Content = T "BtnLoadDBs"
    $Script:ui_btnSelectAll.Content = T "BtnSelectAll"
    $Script:ui_btnDeselectAll.Content = T "BtnDeselectAll"
    $Script:ui_chkHideInstance.Content = T "LblHideInstanceDB"
    $Script:ui_txtExportFolderLabel.Text = T "LblExportFolder"
    $Script:ui_btnExportCsv.Content = T "BtnExportCsv"

    # Export folder default text
    if (-not (Get-AppState).ExportFolder) {
        $Script:ui_txtExportFolder.Text = T "LblDefaultFolder"
    }

    # Refresh DB count
    $filtered = (Get-AppState).FilteredDataBlocks
    $Script:ui_txtDBCount.Text = (T "LblDBCount") -f $filtered.Count

    # Refresh DB list labels (type badges) if blocks loaded
    if ($filtered.Count -gt 0) {
        Refresh-DataBlockList
    }

    # Update tooltips
    $Script:ui_btnScan.ToolTip = T "TipScan"
    $Script:ui_btnConnect.ToolTip = T "TipConnect"
    $Script:ui_btnDisconnect.ToolTip = T "TipDisconnect"
    $Script:ui_btnLoadDBs.ToolTip = T "TipLoadDBs"
    $Script:ui_btnExportCsv.ToolTip = T "TipExportCsv"
    $Script:ui_btnBrowseFolder.ToolTip = T "TipBrowse"
}

# =================== NAVIGATION EVENTS ===================

function Register-NavigationEvents {
    $brush = [System.Windows.Media.BrushConverter]::new()

    $Script:ui_btnNavConnection.Add_Click({
        $Script:ui_pnlConnection.Visibility = [System.Windows.Visibility]::Visible
        $Script:ui_pnlExport.Visibility = [System.Windows.Visibility]::Collapsed
        # Highlight active nav
        $b = [System.Windows.Media.BrushConverter]::new()
        $Script:ui_btnNavConnection.Background = $b.ConvertFrom("#EAF2F8")
        $Script:ui_txtNavConnection.Foreground = $b.ConvertFrom("#1A5276")
        $Script:ui_txtNavConnection.FontWeight = [System.Windows.FontWeights]::SemiBold
        $Script:ui_btnNavExport.Background = [System.Windows.Media.Brushes]::Transparent
        $Script:ui_txtNavExport.Foreground = $b.ConvertFrom("#4A5568")
        $Script:ui_txtNavExport.FontWeight = [System.Windows.FontWeights]::Normal
    })

    $Script:ui_btnNavExport.Add_Click({
        $Script:ui_pnlConnection.Visibility = [System.Windows.Visibility]::Collapsed
        $Script:ui_pnlExport.Visibility = [System.Windows.Visibility]::Visible
        # Highlight active nav
        $b = [System.Windows.Media.BrushConverter]::new()
        $Script:ui_btnNavExport.Background = $b.ConvertFrom("#EAF2F8")
        $Script:ui_txtNavExport.Foreground = $b.ConvertFrom("#1A5276")
        $Script:ui_txtNavExport.FontWeight = [System.Windows.FontWeights]::SemiBold
        $Script:ui_btnNavConnection.Background = [System.Windows.Media.Brushes]::Transparent
        $Script:ui_txtNavConnection.Foreground = $b.ConvertFrom("#4A5568")
        $Script:ui_txtNavConnection.FontWeight = [System.Windows.FontWeights]::Normal
    })
}

# =================== CONNECTION EVENTS ===================

function Register-ConnectionEvents {
    # Scan
    $Script:ui_btnScan.Add_Click({
        try {
            $Script:ui_btnScan.IsEnabled = $false
            $Script:ui_lbInstances.Items.Clear()
            $instances = Invoke-TiaScan
            foreach ($inst in $instances) {
                $item = New-InstanceListItem -Instance $inst
                $Script:ui_lbInstances.Items.Add($item) | Out-Null
            }
            if ($instances.Count -gt 0) {
                Set-StatusBanner -Banner $Script:ui_brdScanStatus -TextBlock $Script:ui_txtScanStatus `
                    -Text ((T "LblScanResult") -f $instances.Count) -Type "info"
            } else {
                Set-StatusBanner -Banner $Script:ui_brdScanStatus -TextBlock $Script:ui_txtScanStatus `
                    -Text (T "LblNoInstance") -Type "warning"
            }
        } catch {
            [System.Windows.MessageBox]::Show(
                ((T "MsgScanError") -f $_.Exception.Message),
                (T "MsgError"), "OK", "Error")
        } finally {
            $Script:ui_btnScan.IsEnabled = $true
        }
    })

    # Connect
    $Script:ui_btnConnect.Add_Click({
        $selectedIndex = $Script:ui_lbInstances.SelectedIndex
        if ($selectedIndex -lt 0) {
            [System.Windows.MessageBox]::Show((T "MsgNoInstance"), (T "MsgInfo"), "OK", "Information")
            return
        }
        $instances = (Get-AppState).TiaInstances
        $selectedInst = $instances[$selectedIndex]

        try {
            $Script:ui_btnConnect.IsEnabled = $false
            $Script:ui_Window.Cursor = [System.Windows.Input.Cursors]::Wait

            $result = Connect-TiaInstance -ProcessId $selectedInst.ProcessId
            if ($result.Success) {
                # Update connection UI
                $Script:ui_brdConnected.Visibility = [System.Windows.Visibility]::Visible
                $Script:ui_brdDisconnected.Visibility = [System.Windows.Visibility]::Collapsed
                $Script:ui_btnDisconnect.IsEnabled = $true
                # Lock version selector
                $Script:ui_cbTiaVersion.IsEnabled = $false
                $Script:ui_txtVersionInfo.Text = T "LblVersionLocked"

                [System.Windows.MessageBox]::Show($result.Message, (T "MsgInfo"), "OK", "Information")
            } else {
                [System.Windows.MessageBox]::Show($result.Message, (T "MsgError"), "OK", "Warning")
            }
        } catch {
            [System.Windows.MessageBox]::Show(
                ((T "MsgConnectError") -f $_.Exception.Message),
                (T "MsgError"), "OK", "Error")
        } finally {
            $Script:ui_btnConnect.IsEnabled = $true
            $Script:ui_Window.Cursor = $null
        }
    })

    # Disconnect
    $Script:ui_btnDisconnect.Add_Click({
        Disconnect-TiaInstance
        $Script:ui_brdConnected.Visibility = [System.Windows.Visibility]::Collapsed
        $Script:ui_brdDisconnected.Visibility = [System.Windows.Visibility]::Visible
        $Script:ui_btnDisconnect.IsEnabled = $false
        $Script:ui_cbTiaVersion.IsEnabled = $true
        $Script:ui_lbDataBlocks.Items.Clear()
        $Script:ui_txtDBCount.Text = (T "LblDBCount") -f 0

        $versions = (Get-AppState).InstalledVersions
        $match = $versions | Where-Object { $_.Version -eq $Script:ui_cbTiaVersion.SelectedItem }
        if ($match) {
            $Script:ui_txtVersionInfo.Text = (T "LblDllLoaded") -f [System.IO.Path]::GetFileName($match.DllPath)
        }
    })
}

# =================== EXPORT EVENTS ===================

function Register-ExportEvents {
    # Load DataBlocks
    $Script:ui_btnLoadDBs.Add_Click({
        if (-not (Get-AppState).IsConnected) {
            [System.Windows.MessageBox]::Show((T "MsgConnectFirst"), (T "MsgInfo"), "OK", "Information")
            return
        }
        try {
            $Script:ui_btnLoadDBs.IsEnabled = $false
            $Script:ui_Window.Cursor = [System.Windows.Input.Cursors]::Wait
            Get-AllDataBlocks
            Refresh-DataBlockList
            Refresh-PlcInfoPanel
        } catch {
            [System.Windows.MessageBox]::Show(
                ((T "MsgLoadError") -f $_.Exception.Message),
                (T "MsgError"), "OK", "Error")
        } finally {
            $Script:ui_btnLoadDBs.IsEnabled = $true
            $Script:ui_Window.Cursor = $null
        }
    })

    # Select All / Deselect All
    $Script:ui_btnSelectAll.Add_Click({
        Set-AllBlocksSelected -Selected $true
    })
    $Script:ui_btnDeselectAll.Add_Click({
        Set-AllBlocksSelected -Selected $false
    })

    # Hide Instance DBs toggle
    $Script:ui_chkHideInstance.Add_Checked({
        Set-AppStateValue -Key "HideInstanceDBs" -Value $true
        Apply-DataBlockFilter
        Refresh-DataBlockList
    })
    $Script:ui_chkHideInstance.Add_Unchecked({
        Set-AppStateValue -Key "HideInstanceDBs" -Value $false
        Apply-DataBlockFilter
        Refresh-DataBlockList
    })

    # Browse folder
    $Script:ui_btnBrowseFolder.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = T "TipBrowse"
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Set-AppStateValue -Key "ExportFolder" -Value $dialog.SelectedPath
            $Script:ui_txtExportFolder.Text = $dialog.SelectedPath
        }
    })

    # Export CSV
    $Script:ui_btnExportCsv.Add_Click({
        $state = Get-AppState
        $selected = @($state.FilteredDataBlocks | Where-Object { $_.IsSelected })

        if ($selected.Count -eq 0) {
            [System.Windows.MessageBox]::Show((T "MsgNoSelection"), (T "MsgInfo"), "OK", "Information")
            return
        }

        try {
            Set-AppStateValue -Key "IsExporting" -Value $true
            $Script:ui_btnExportCsv.IsEnabled = $false
            $Script:ui_Window.Cursor = [System.Windows.Input.Cursors]::Wait

            $folder = New-ExportFolder -BasePath $state.ExportFolder
            $result = Invoke-TableExport -SelectedBlocks $selected -OutputFolder $folder
            $summary = Get-ExportSummary -Result $result
            $icon = if ($result.ErrorCount -gt 0 -or $result.OptimizedDBs.Count -gt 0) { "Warning" } else { "Information" }
            [System.Windows.MessageBox]::Show($summary, (T "MsgExportCsvDone"), "OK", $icon)
        } catch {
            [System.Windows.MessageBox]::Show(
                ((T "MsgLoadError") -f $_.Exception.Message),
                (T "MsgError"), "OK", "Error")
        } finally {
            Set-AppStateValue -Key "IsExporting" -Value $false
            $Script:ui_btnExportCsv.IsEnabled = $true
            $Script:ui_Window.Cursor = $null
        }
    })
}

# =================== DATA BLOCK LIST ===================

function Refresh-DataBlockList {
    $Script:ui_lbDataBlocks.Items.Clear()
    $blocks = (Get-AppState).FilteredDataBlocks
    $Script:ui_txtDBCount.Text = (T "LblDBCount") -f $blocks.Count

    foreach ($db in $blocks) {
        # Update display type label with current language
        $db.DisplayType = if ($db.IsInstanceDB) { T "TypeInstance" } else { T "TypeGlobal" }
        $item = New-DataBlockListItem -Block $db
        $Script:ui_lbDataBlocks.Items.Add($item) | Out-Null
    }
}

function Set-AllBlocksSelected {
    param([bool]$Selected)
    $blocks = (Get-AppState).FilteredDataBlocks
    foreach ($b in $blocks) { $b.IsSelected = $Selected }
    Refresh-DataBlockList
}

# =================== PLC INFO PANEL ===================

function Refresh-PlcInfoPanel {
    $Script:ui_spPlcInfo.Children.Clear()
    $plcInfoList = (Get-AppState).PlcDeviceInfoList

    if ($plcInfoList.Count -eq 0) {
        $Script:ui_brdPlcInfo.Visibility = [System.Windows.Visibility]::Collapsed
        return
    }

    $Script:ui_brdPlcInfo.Visibility = [System.Windows.Visibility]::Visible
    $brush = [System.Windows.Media.BrushConverter]::new()

    foreach ($plcInfo in $plcInfoList) {
        $panel = [System.Windows.Controls.StackPanel]::new()
        $panel.Orientation = [System.Windows.Controls.Orientation]::Horizontal
        $panel.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)

        # PLC name
        $txtName = [System.Windows.Controls.TextBlock]::new()
        $txtName.Text = (T "LblPlcName") -f $plcInfo.Name
        $txtName.FontWeight = [System.Windows.FontWeights]::SemiBold
        $txtName.FontSize = 12
        $txtName.Margin = [System.Windows.Thickness]::new(0, 0, 16, 0)
        $txtName.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $panel.Children.Add($txtName) | Out-Null

        # IP
        $txtIp = [System.Windows.Controls.TextBlock]::new()
        $ipText = if ($plcInfo.IpAddress) { $plcInfo.IpAddress } else { "N/A" }
        $txtIp.Text = (T "LblPlcIp") -f $ipText
        $txtIp.FontSize = 11
        $txtIp.Foreground = $brush.ConvertFrom("#4A5568")
        $txtIp.Margin = [System.Windows.Thickness]::new(0, 0, 16, 0)
        $txtIp.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $panel.Children.Add($txtIp) | Out-Null

        # TSAP
        $txtTsap = [System.Windows.Controls.TextBlock]::new()
        $txtTsap.Text = (T "LblPlcTsap") -f $plcInfo.Tsap
        $txtTsap.FontSize = 11
        $txtTsap.Foreground = $brush.ConvertFrom("#4A5568")
        $txtTsap.Margin = [System.Windows.Thickness]::new(0, 0, 16, 0)
        $txtTsap.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $panel.Children.Add($txtTsap) | Out-Null

        # Edit button
        $btnEdit = [System.Windows.Controls.Button]::new()
        $btnEdit.Content = T "BtnEditPlcInfo"
        $btnEdit.FontSize = 10
        $btnEdit.Padding = [System.Windows.Thickness]::new(8, 2, 8, 2)
        $btnEdit.Cursor = [System.Windows.Input.Cursors]::Hand
        $btnEdit.Background = $brush.ConvertFrom("White")
        $btnEdit.BorderBrush = $brush.ConvertFrom("#CBD5E0")
        $btnEdit.Tag = $plcInfo.PlcIndex
        $btnEdit.Add_Click({
            param($sender, $e)
            $plcIdx = $sender.Tag
            Show-PlcInfoEditDialog -PlcIndex $plcIdx
            Refresh-PlcInfoPanel
        })
        $panel.Children.Add($btnEdit) | Out-Null

        $Script:ui_spPlcInfo.Children.Add($panel) | Out-Null
    }
}

function Show-PlcInfoEditDialog {
    param([int]$PlcIndex)

    $plcInfoList = (Get-AppState).PlcDeviceInfoList
    $plcInfo = $plcInfoList | Where-Object { $_.PlcIndex -eq $PlcIndex } | Select-Object -First 1
    if (-not $plcInfo) { return }

    # Create a simple edit dialog programmatically
    $dlg = [System.Windows.Window]::new()
    $dlg.Title = T "DlgPlcInfoTitle"
    $dlg.Width = 380
    $dlg.Height = 300
    $dlg.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
    $dlg.Owner = $Script:ui_Window
    $dlg.ResizeMode = [System.Windows.ResizeMode]::NoResize

    $stack = [System.Windows.Controls.StackPanel]::new()
    $stack.Margin = [System.Windows.Thickness]::new(16)

    # PLC Name field
    $lblName = [System.Windows.Controls.TextBlock]::new()
    $lblName.Text = T "DlgPlcName"
    $lblName.FontSize = 12
    $lblName.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)
    $stack.Children.Add($lblName) | Out-Null

    $txtName = [System.Windows.Controls.TextBox]::new()
    $txtName.Text = $plcInfo.Name
    $txtName.FontSize = 13
    $txtName.Padding = [System.Windows.Thickness]::new(6, 4, 6, 4)
    $txtName.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
    $stack.Children.Add($txtName) | Out-Null

    # IP Address field
    $lblIp = [System.Windows.Controls.TextBlock]::new()
    $lblIp.Text = T "DlgPlcIp"
    $lblIp.FontSize = 12
    $lblIp.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)
    $stack.Children.Add($lblIp) | Out-Null

    $txtIp = [System.Windows.Controls.TextBox]::new()
    $txtIp.Text = $plcInfo.IpAddress
    $txtIp.FontSize = 13
    $txtIp.Padding = [System.Windows.Thickness]::new(6, 4, 6, 4)
    $txtIp.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
    $stack.Children.Add($txtIp) | Out-Null

    # TSAP field
    $lblTsap = [System.Windows.Controls.TextBlock]::new()
    $lblTsap.Text = T "DlgPlcTsap"
    $lblTsap.FontSize = 12
    $lblTsap.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)
    $stack.Children.Add($lblTsap) | Out-Null

    $txtTsap = [System.Windows.Controls.TextBox]::new()
    $txtTsap.Text = $plcInfo.Tsap
    $txtTsap.FontSize = 13
    $txtTsap.Padding = [System.Windows.Thickness]::new(6, 4, 6, 4)
    $txtTsap.Margin = [System.Windows.Thickness]::new(0, 0, 0, 16)
    $stack.Children.Add($txtTsap) | Out-Null

    # OK button
    $btnOk = [System.Windows.Controls.Button]::new()
    $btnOk.Content = "OK"
    $btnOk.Width = 100
    $btnOk.Height = 32
    $btnOk.FontSize = 13
    $btnOk.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
    $btnOk.IsDefault = $true
    $btnOk.Add_Click({
        $plcInfo.Name = $txtName.Text
        $plcInfo.IpAddress = $txtIp.Text
        $plcInfo.Tsap = $txtTsap.Text
        $dlg.Close()
    }.GetNewClosure())
    $stack.Children.Add($btnOk) | Out-Null

    $dlg.Content = $stack
    $dlg.ShowDialog() | Out-Null
}
