# =============================================================================================================
# Script:    HyperVExplorer.ps1
# Version:   1.3
# Purpose:   WPF GUI tool for remote Hyper-V inventory collection across multiple hosts
# Requires:  PowerShell 5.1+, WinRM enabled on target Hyper-V hosts
# =============================================================================================================


Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ---- Config file management ----
function Get-ConfigPath {
    $dir = Join-Path $env:APPDATA "HyperVExplorer"
    if (!(Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Join-Path $dir "config.json"
}

function Load-Config {
    $path = Get-ConfigPath
    if (Test-Path $path) {
        try {
            return (Get-Content -Path $path -Raw | ConvertFrom-Json)
        } catch {
            return [PSCustomObject]@{ version = 1; hosts = @() }
        }
    }
    return [PSCustomObject]@{ version = 1; hosts = @() }
}

function Save-Config {
    param($Config)
    $path = Get-ConfigPath
    $Config | ConvertTo-Json -Depth 4 | Set-Content -Path $path -Encoding UTF8
}

function Save-HostToHistory {
    param(
        [string]$Address,
        [bool]$UseCurrentUser,
        [System.Management.Automation.PSCredential]$Credential,
        [bool]$RememberCredential
    )
    $config = Load-Config
    $existingHosts = @($config.hosts | Where-Object { $_.address -ne $Address })

    $entry = [PSCustomObject]@{
        address           = $Address
        lastConnected     = (Get-Date).ToString("o")
        useCurrentUser    = $UseCurrentUser
        username          = $null
        encryptedPassword = $null
    }

    if ($Credential -and $RememberCredential) {
        $entry.username = $Credential.UserName
        $entry.encryptedPassword = ($Credential.Password | ConvertFrom-SecureString)
    }

    $allHosts = @($entry) + $existingHosts
    if ($allHosts.Count -gt 20) { $allHosts = $allHosts[0..19] }
    $config.hosts = $allHosts

    Save-Config $config
}

function Get-SavedCredential {
    param([string]$Address)
    $config = Load-Config
    $entry = $config.hosts | Where-Object { $_.address -eq $Address } | Select-Object -First 1
    if ($entry -and $entry.username -and $entry.encryptedPassword) {
        try {
            $secPass = $entry.encryptedPassword | ConvertTo-SecureString
            return [System.Management.Automation.PSCredential]::new($entry.username, $secPass)
        } catch {
            return $null
        }
    }
    return $null
}

function Get-HostHistoryEntry {
    param([string]$Address)
    $config = Load-Config
    return ($config.hosts | Where-Object { $_.address -eq $Address } | Select-Object -First 1)
}

# ---- XAML UI Definition ----
[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="HyperV Explorer"
    Width="1400" Height="750"
    MinWidth="900" MinHeight="500"
    WindowStartupLocation="CenterScreen"
    Background="#1e1e2e">

    <Window.Resources>
        <!-- Dark theme styles -->
        <SolidColorBrush x:Key="BgDark" Color="#1e1e2e"/>
        <SolidColorBrush x:Key="BgMedium" Color="#313244"/>
        <SolidColorBrush x:Key="BgLight" Color="#45475a"/>
        <SolidColorBrush x:Key="FgText" Color="#cdd6f4"/>
        <SolidColorBrush x:Key="FgSubtle" Color="#a6adc8"/>
        <SolidColorBrush x:Key="AccentBlue" Color="#89b4fa"/>
        <SolidColorBrush x:Key="AccentGreen" Color="#a6e3a1"/>
        <SolidColorBrush x:Key="AccentRed" Color="#f38ba8"/>
        <SolidColorBrush x:Key="AccentYellow" Color="#f9e2af"/>

        <Style TargetType="Button">
            <Setter Property="Background" Value="{StaticResource BgLight}"/>
            <Setter Property="Foreground" Value="{StaticResource FgText}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#585b70"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{StaticResource BgMedium}"/>
            <Setter Property="Foreground" Value="{StaticResource FgText}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BgLight}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="CaretBrush" Value="{StaticResource FgText}"/>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource FgText}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title Bar -->
        <Border Grid.Row="0" Background="{StaticResource BgMedium}" Padding="16,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="&#xE7F4; HyperV Explorer" FontSize="20" FontWeight="SemiBold"
                           Foreground="{StaticResource AccentBlue}" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="2" Text="v1.3" FontSize="12"
                           Foreground="{StaticResource FgSubtle}" VerticalAlignment="Center"/>
            </Grid>
        </Border>

        <!-- Connection Toolbar -->
        <Border Grid.Row="1" Background="{StaticResource BgDark}" Padding="16,10">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                <TextBlock Text="Host:" Foreground="{StaticResource FgSubtle}" VerticalAlignment="Center"
                           FontSize="13" Margin="0,0,8,0"/>
                <TextBox x:Name="txtHost" Width="230" VerticalAlignment="Center"
                         ToolTip="Enter hostname or IP address"/>
                <Button x:Name="btnHistory" Content="&#x25BE;" Margin="2,0,0,0" Padding="8,8"
                        Background="{StaticResource BgMedium}" Foreground="{StaticResource FgSubtle}"
                        ToolTip="Recent hosts" FontSize="13"/>
                <CheckBox x:Name="chkCurrentUser" Content="Use current user" IsChecked="True"
                          Margin="16,0,0,0" VerticalAlignment="Center"/>
                <Button x:Name="btnConnect" Content="&#x2B; Connect" Margin="16,0,0,0"
                        Background="#364a63" Foreground="{StaticResource AccentBlue}"/>
                <Button x:Name="btnBulkConnect" Content="&#x21C4; Bulk Connect" Margin="8,0,0,0"
                        Background="#364a63" Foreground="{StaticResource AccentBlue}"
                        ToolTip="Connect to multiple saved hosts at once"/>
                <Button x:Name="btnDisconnect" Content="&#x2716; Disconnect Selected Host" Margin="8,0,0,0"
                        Background="#4a3644" Foreground="{StaticResource AccentRed}" IsEnabled="False"/>
                <Rectangle Width="1" Fill="{StaticResource BgLight}" Margin="16,2" VerticalAlignment="Stretch"/>
                <Button x:Name="btnExport" Content="&#x1F4BE; Export CSV" Margin="0,0,0,0"
                        Background="#3a4a3a" Foreground="{StaticResource AccentGreen}" IsEnabled="False"/>
                <Button x:Name="btnClear" Content="Clear All" Margin="8,0,0,0" IsEnabled="False"/>
            </StackPanel>
        </Border>

        <!-- DataGrid -->
        <DataGrid Grid.Row="2" x:Name="dgVMs" Margin="16,8,16,8"
                  AutoGenerateColumns="False" IsReadOnly="True"
                  CanUserSortColumns="True" CanUserReorderColumns="True" CanUserResizeColumns="True"
                  SelectionMode="Extended" SelectionUnit="FullRow"
                  GridLinesVisibility="Horizontal"
                  Background="{StaticResource BgMedium}"
                  Foreground="{StaticResource FgText}"
                  BorderBrush="{StaticResource BgLight}"
                  BorderThickness="1"
                  RowBackground="#2a2a3c"
                  AlternatingRowBackground="#313244"
                  HorizontalGridLinesBrush="#3a3a4c"
                  HeadersVisibility="Column"
                  FontSize="12">

            <DataGrid.ColumnHeaderStyle>
                <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Background" Value="#3a3a5c"/>
                    <Setter Property="Foreground" Value="{StaticResource AccentBlue}"/>
                    <Setter Property="Padding" Value="8,6"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="FontSize" Value="12"/>
                    <Setter Property="BorderBrush" Value="#4a4a6c"/>
                    <Setter Property="BorderThickness" Value="0,0,1,1"/>
                </Style>
            </DataGrid.ColumnHeaderStyle>

            <DataGrid.Columns>
                <DataGridTextColumn Header="Host" Binding="{Binding HostName}" Width="120"/>
                <DataGridTextColumn Header="Host CPU" Binding="{Binding HostCPU}" Width="70"/>
                <DataGridTextColumn Header="Host Mem (GB)" Binding="{Binding HostMemoryGB}" Width="95"/>
                <DataGridTextColumn Header="Host Ver" Binding="{Binding HostVersion}" Width="80"/>
                <DataGridTextColumn Header="VM Name" Binding="{Binding VMName}" Width="160"/>
                <DataGridTextColumn Header="State" Binding="{Binding State}" Width="75"/>
                <DataGridTextColumn Header="vCPU" Binding="{Binding CPUCount}" Width="55"/>
                <DataGridTextColumn Header="Mem (MB)" Binding="{Binding MemoryAssignedMB}" Width="80"/>
                <DataGridTextColumn Header="Uptime" Binding="{Binding Uptime}" Width="100"/>
                <DataGridTextColumn Header="Gen" Binding="{Binding Generation}" Width="45"/>
                <DataGridTextColumn Header="Dyn Mem" Binding="{Binding DynamicMemory}" Width="70"/>
                <DataGridTextColumn Header="NICs" Binding="{Binding NICs}" Width="250"/>
                <DataGridTextColumn Header="Disks" Binding="{Binding Disks}" Width="350"/>
                <DataGridTextColumn Header="Checkpoints" Binding="{Binding Checkpoints}" Width="150"/>
                <DataGridTextColumn Header="Int. Services" Binding="{Binding IntegrationServices}" Width="100"/>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Status Bar -->
        <Border Grid.Row="3" Background="{StaticResource BgMedium}" Padding="16,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="txtStatus" Grid.Column="0" Text="Ready — enter a host and click Connect"
                           Foreground="{StaticResource FgSubtle}" FontSize="12" VerticalAlignment="Center"/>
                <ProgressBar x:Name="pbProgress" Grid.Column="2" Width="120" Height="8"
                             IsIndeterminate="False" Visibility="Collapsed" Margin="0,0,16,0"
                             Background="{StaticResource BgLight}" Foreground="{StaticResource AccentBlue}"/>
                <TextBlock x:Name="txtVMCount" Grid.Column="3" Text="0 VMs | 0 Hosts"
                           Foreground="{StaticResource FgSubtle}" FontSize="12" VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ---- Load XAML and find controls ----
$Reader = [System.Xml.XmlNodeReader]::new($XAML)
$Window = [Windows.Markup.XamlReader]::Load($Reader)

$txtHost         = $Window.FindName("txtHost")
$chkCurrentUser  = $Window.FindName("chkCurrentUser")
$btnConnect      = $Window.FindName("btnConnect")
$btnBulkConnect  = $Window.FindName("btnBulkConnect")
$btnDisconnect   = $Window.FindName("btnDisconnect")
$btnExport       = $Window.FindName("btnExport")
$btnClear        = $Window.FindName("btnClear")
$btnHistory      = $Window.FindName("btnHistory")
$dgVMs           = $Window.FindName("dgVMs")
$txtStatus       = $Window.FindName("txtStatus")
$pbProgress      = $Window.FindName("pbProgress")
$txtVMCount      = $Window.FindName("txtVMCount")

# ---- State ----
$script:VMData = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new()
$dgVMs.ItemsSource = $script:VMData
$script:ConnectedHosts = @{}  # HostName -> @{ Credential = $cred; VMCount = N }
$script:AppConfig = Load-Config

# ---- Check local WinRM service at startup ----
$WinRMService = Get-Service -Name WinRM -ErrorAction SilentlyContinue
if ($WinRMService -and $WinRMService.Status -ne 'Running') {
    $StartIt = [System.Windows.MessageBox]::Show(
        "The WinRM service is not running on this machine.`n`nIt is required for remote connections. Start it now?",
        "WinRM Service Required",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)
    if ($StartIt -eq 'Yes') {
        try {
            Start-Service -Name WinRM -ErrorAction Stop
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Could not start WinRM service.`n`nError: $($_.Exception.Message)`n`nTry running as Administrator.",
                "Service Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error)
        }
    }
}

# ---- Helper functions ----
function Update-StatusBar {
    $hostCount = $script:ConnectedHosts.Count
    $vmCount   = $script:VMData.Count
    $txtVMCount.Text = "$vmCount VMs | $hostCount Hosts"
    $hasData = $vmCount -gt 0
    $btnExport.IsEnabled     = $hasData
    $btnClear.IsEnabled      = $hasData
    $btnDisconnect.IsEnabled = $hasData
}

function Set-Status {
    param([string]$Message, [string]$Color = "#a6adc8")
    $txtStatus.Text = $Message
    $txtStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
    # Force UI refresh
    $Window.Dispatcher.Invoke([Action]{}, 'Background')
}

function Show-Progress {
    param([bool]$Show)
    $pbProgress.IsIndeterminate = $Show
    $pbProgress.Visibility = if ($Show) { "Visible" } else { "Collapsed" }
    $Window.Dispatcher.Invoke([Action]{}, 'Background')
}

function Test-IsIPAddress {
    param([string]$Value)
    $ip = $null
    [System.Net.IPAddress]::TryParse($Value, [ref]$ip)
}

function Show-CredentialDialog {
    param(
        [string]$TargetHost,
        [string]$PreFillUser = "",
        [bool]$ShowRemember = $true
    )

    [xml]$CredXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Credentials" Width="420" Height="270"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        Background="#1e1e2e">
    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="Enter credentials for $TargetHost"
                   Foreground="#cdd6f4" FontSize="14" Margin="0,0,0,16"/>
        <Grid Grid.Row="1" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="90"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Username:" Foreground="#a6adc8" VerticalAlignment="Center" FontSize="13"/>
            <TextBox x:Name="txtUser" Grid.Column="1" FontSize="13"
                     Background="#313244" Foreground="#cdd6f4" BorderBrush="#45475a"
                     Padding="8,6" CaretBrush="#cdd6f4"/>
        </Grid>
        <Grid Grid.Row="2" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="90"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Password:" Foreground="#a6adc8" VerticalAlignment="Center" FontSize="13"/>
            <PasswordBox x:Name="txtPass" Grid.Column="1" FontSize="13"
                         Background="#313244" Foreground="#cdd6f4" BorderBrush="#45475a"
                         Padding="8,6" CaretBrush="#cdd6f4"/>
        </Grid>
        <ContentControl Grid.Row="3" Margin="0,0,0,10" x:Name="rememberPlaceholder"/>
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnOK" Content="Connect" Width="90" Padding="8,6" Margin="0,0,8,0"
                    Background="#364a63" Foreground="#89b4fa" FontSize="13" IsDefault="True"/>
            <Button x:Name="btnCancel" Content="Cancel" Width="90" Padding="8,6"
                    Background="#45475a" Foreground="#cdd6f4" FontSize="13" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $CredReader = [System.Xml.XmlNodeReader]::new($CredXAML)
    $CredWindow = [Windows.Markup.XamlReader]::Load($CredReader)
    $CredWindow.Owner = $Window

    $credTxtUser  = $CredWindow.FindName("txtUser")
    $credTxtPass  = $CredWindow.FindName("txtPass")
    $credBtnOK    = $CredWindow.FindName("btnOK")
    $credBtnCancel = $CredWindow.FindName("btnCancel")
    $rememberPlaceholder = $CredWindow.FindName("rememberPlaceholder")

    # Add Remember checkbox dynamically
    $chkRemember = $null
    if ($ShowRemember) {
        $chkRemember = [System.Windows.Controls.CheckBox]::new()
        $chkRemember.Content = "Remember credentials for this host"
        $chkRemember.IsChecked = $true
        $chkRemember.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#a6adc8")
        $chkRemember.FontSize = 12
        $rememberPlaceholder.Content = $chkRemember
    }

    # Pre-fill username if available
    if ($PreFillUser) {
        $credTxtUser.Text = $PreFillUser
        $credTxtPass.Focus() | Out-Null
    } else {
        $credTxtUser.Focus() | Out-Null
    }

    $credBtnOK.Add_Click({
        if ([string]::IsNullOrWhiteSpace($credTxtUser.Text)) {
            [System.Windows.MessageBox]::Show("Please enter a username.", "Credentials",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        $CredWindow.DialogResult = $true
    }.GetNewClosure())

    $credBtnCancel.Add_Click({
        $CredWindow.DialogResult = $false
    }.GetNewClosure())

    $dialogResult = $CredWindow.ShowDialog()

    if ($dialogResult -eq $true) {
        $user = $credTxtUser.Text.Trim()
        $pass = $credTxtPass.SecurePassword
        $remember = if ($chkRemember) { $chkRemember.IsChecked -eq $true } else { $false }
        return @{
            Credential = [System.Management.Automation.PSCredential]::new($user, $pass)
            Remember   = $remember
        }
    }
    return $null
}

function Show-BulkConnectDialog {
    $config = Load-Config
    $hostList = @($config.hosts)

    if ($hostList.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No saved hosts found.`n`nConnect to hosts individually first to build your history.",
            "Bulk Connect",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information)
        return $null
    }

    [xml]$BulkXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Bulk Connect — Select Hosts" Width="550" Height="450"
        WindowStartupLocation="CenterOwner" ResizeMode="CanResize"
        MinWidth="400" MinHeight="300"
        Background="#1e1e2e">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Select hosts to connect to:" Foreground="#cdd6f4"
                   FontSize="14" Margin="0,0,0,12"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <Button x:Name="btnSelectAll" Content="Select All" Padding="12,4" FontSize="12"
                    Background="#45475a" Foreground="#cdd6f4" Margin="0,0,8,0"/>
            <Button x:Name="btnSelectNone" Content="Select None" Padding="12,4" FontSize="12"
                    Background="#45475a" Foreground="#cdd6f4"/>
        </StackPanel>
        <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto" Background="#313244"
                      Padding="8">
            <StackPanel x:Name="hostListPanel"/>
        </ScrollViewer>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnBulkOK" Content="Connect Selected" Width="140" Padding="8,6"
                    Margin="0,0,8,0" Background="#364a63" Foreground="#89b4fa" FontSize="13"/>
            <Button x:Name="btnBulkCancel" Content="Cancel" Width="90" Padding="8,6"
                    Background="#45475a" Foreground="#cdd6f4" FontSize="13" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $BulkReader = [System.Xml.XmlNodeReader]::new($BulkXAML)
    $BulkWindow = [Windows.Markup.XamlReader]::Load($BulkReader)
    $BulkWindow.Owner = $Window

    $hostListPanel = $BulkWindow.FindName("hostListPanel")
    $btnSelectAll  = $BulkWindow.FindName("btnSelectAll")
    $btnSelectNone = $BulkWindow.FindName("btnSelectNone")
    $btnBulkOK     = $BulkWindow.FindName("btnBulkOK")
    $btnBulkCancel = $BulkWindow.FindName("btnBulkCancel")

    # Build checkbox list for each saved host
    $checkboxes = @()
    foreach ($h in $hostList) {
        $authInfo = if ($h.useCurrentUser) { "Current User" }
                    elseif ($h.username) { "Saved: $($h.username)" }
                    else { "Will prompt" }
        $lastDate = try { ([datetime]$h.lastConnected).ToString("yyyy-MM-dd HH:mm") } catch { "Unknown" }
        $alreadyConnected = $script:ConnectedHosts.ContainsKey($h.address)

        $cb = [System.Windows.Controls.CheckBox]::new()
        $cb.IsChecked = (-not $alreadyConnected)
        $cb.IsEnabled = (-not $alreadyConnected)
        $cb.Margin = [System.Windows.Thickness]::new(0, 4, 0, 4)
        $cb.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#cdd6f4")
        $cb.FontSize = 13
        $cb.Tag = $h.address

        $label = "$($h.address)  —  $authInfo  —  Last: $lastDate"
        if ($alreadyConnected) { $label += "  (already connected)" }
        $cb.Content = $label

        $hostListPanel.Children.Add($cb) | Out-Null
        $checkboxes += $cb
    }

    $btnSelectAll.Add_Click({
        foreach ($c in $checkboxes) { if ($c.IsEnabled) { $c.IsChecked = $true } }
    }.GetNewClosure())

    $btnSelectNone.Add_Click({
        foreach ($c in $checkboxes) { if ($c.IsEnabled) { $c.IsChecked = $false } }
    }.GetNewClosure())

    $btnBulkOK.Add_Click({
        $BulkWindow.DialogResult = $true
    }.GetNewClosure())

    $btnBulkCancel.Add_Click({
        $BulkWindow.DialogResult = $false
    }.GetNewClosure())

    $dialogResult = $BulkWindow.ShowDialog()

    if ($dialogResult -eq $true) {
        $selected = @()
        foreach ($c in $checkboxes) {
            if ($c.IsChecked -eq $true -and $c.IsEnabled) {
                $selected += $c.Tag
            }
        }
        return $selected
    }
    return $null
}

function Show-HistoryMenu {
    $config = Load-Config
    $hostList = @($config.hosts)

    $menu = [System.Windows.Controls.ContextMenu]::new()
    $menu.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#313244")
    $menu.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#45475a")
    $menu.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#cdd6f4")

    if ($hostList.Count -eq 0) {
        $emptyItem = [System.Windows.Controls.MenuItem]::new()
        $emptyItem.Header = "No recent hosts"
        $emptyItem.IsEnabled = $false
        $emptyItem.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#a6adc8")
        $menu.Items.Add($emptyItem) | Out-Null
    } else {
        foreach ($h in $hostList) {
            $authTag = if ($h.useCurrentUser) { "[Current User]" }
                       elseif ($h.username) { "[$($h.username)]" }
                       else { "[No saved creds]" }
            $lastDate = try { ([datetime]$h.lastConnected).ToString("MM/dd HH:mm") } catch { "" }

            $item = [System.Windows.Controls.MenuItem]::new()
            $item.Header = "$($h.address)  $authTag  $lastDate"
            $item.Tag = $h.address
            $item.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#cdd6f4")
            $item.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#313244")

            $hostAddress = $h.address
            $hostUseCurrentUser = [bool]$h.useCurrentUser
            $item.Add_Click({
                $txtHost.Text = $hostAddress
                $chkCurrentUser.IsChecked = $hostUseCurrentUser
            }.GetNewClosure())

            $menu.Items.Add($item) | Out-Null
        }

        # Separator + Clear History
        $menu.Items.Add([System.Windows.Controls.Separator]::new()) | Out-Null

        $clearItem = [System.Windows.Controls.MenuItem]::new()
        $clearItem.Header = "Clear History"
        $clearItem.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#f38ba8")
        $clearItem.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#313244")
        $clearItem.Add_Click({
            $confirm = [System.Windows.MessageBox]::Show(
                "Clear all saved hosts and credentials?",
                "Clear History",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question)
            if ($confirm -eq 'Yes') {
                $config = Load-Config
                $config.hosts = @()
                Save-Config $config
                Set-Status "History cleared." "#f9e2af"
            }
        }.GetNewClosure())
        $menu.Items.Add($clearItem) | Out-Null
    }

    $menu.PlacementTarget = $btnHistory
    $menu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
    $menu.IsOpen = $true
}

# ---- Connect to a single host (reusable logic) ----
function Connect-HyperVHost {
    param(
        [string]$TargetHost,
        [bool]$UseCurrentUser,
        [System.Management.Automation.PSCredential]$ProvidedCredential,
        [bool]$RememberCredential = $false,
        [bool]$SkipPrompts = $false
    )

    if ([string]::IsNullOrWhiteSpace($TargetHost)) { return $false }

    if ($script:ConnectedHosts.ContainsKey($TargetHost)) {
        if (-not $SkipPrompts) {
            [System.Windows.MessageBox]::Show("Host '$TargetHost' is already connected.", "HyperV Explorer",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        return $false
    }

    # Ensure local WinRM is running
    $svc = Get-Service -Name WinRM -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -ne 'Running') {
        try { Start-Service -Name WinRM -ErrorAction Stop }
        catch {
            if (-not $SkipPrompts) {
                [System.Windows.MessageBox]::Show(
                    "WinRM service is not running and could not be started.`n`nError: $($_.Exception.Message)",
                    "WinRM Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
            return $false
        }
    }

    # Pre-connection checks
    Set-Status "Pinging $TargetHost ..." "#89b4fa"
    Show-Progress $true
    $Window.Dispatcher.Invoke([Action]{}, 'Background')

    $PingOK = Test-Connection -ComputerName $TargetHost -Count 2 -Quiet -ErrorAction SilentlyContinue
    if (-not $PingOK) {
        Show-Progress $false
        Set-Status "Host unreachable: $TargetHost" "#f38ba8"
        if (-not $SkipPrompts) {
            [System.Windows.MessageBox]::Show(
                "Cannot reach '$TargetHost'.`n`nPing failed. Please check:`n- The hostname or IP is correct`n- The host is powered on and on the network`n- Firewalls allow ICMP",
                "Host Unreachable", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
        return $false
    }
    Set-Status "Ping OK — checking WinRM on $TargetHost ..." "#89b4fa"
    $Window.Dispatcher.Invoke([Action]{}, 'Background')

    # Detect IP address
    $IsIP = Test-IsIPAddress $TargetHost

    if ($IsIP -and $UseCurrentUser) {
        if (-not $SkipPrompts) {
            $Answer = [System.Windows.MessageBox]::Show(
                "You are connecting by IP address ($($TargetHost)).`n`nKerberos authentication does not work with IP addresses. You need to either:`n`n1. Use a hostname instead of an IP`n2. Provide explicit credentials (click Yes)`n`nWould you like to enter credentials now?",
                "IP Address Detected",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning)
            if ($Answer -ne 'Yes') {
                Show-Progress $false
                return $false
            }
        }
        $UseCurrentUser = $false
    }

    # Build connection parameters
    $InvokeParams = @{
        ComputerName = $TargetHost
        ScriptBlock  = $CollectionScript
        ErrorAction  = 'Stop'
    }

    $Credential = $ProvidedCredential
    $CredRemember = $RememberCredential
    if (-not $UseCurrentUser -and -not $Credential) {
        # Check for saved credentials first
        $SavedCred = Get-SavedCredential -Address $TargetHost
        if ($SavedCred) {
            $UseSaved = [System.Windows.MessageBox]::Show(
                "Use saved credentials for '$TargetHost'?`n`nUser: $($SavedCred.UserName)",
                "Saved Credentials",
                [System.Windows.MessageBoxButton]::YesNoCancel,
                [System.Windows.MessageBoxImage]::Question)
            if ($UseSaved -eq 'Yes') {
                $Credential = $SavedCred
                $CredRemember = $true
            }
            elseif ($UseSaved -eq 'Cancel') {
                Show-Progress $false
                return $false
            }
            # 'No' falls through to prompt for new creds
        }

        if (-not $Credential) {
            $histEntry = Get-HostHistoryEntry -Address $TargetHost
            $preFillUser = if ($histEntry -and $histEntry.username) { $histEntry.username } else { "" }
            $credResult = Show-CredentialDialog -TargetHost $TargetHost -PreFillUser $preFillUser
            if (-not $credResult) {
                Show-Progress $false
                Set-Status "Connection cancelled." "#f9e2af"
                return $false
            }
            $Credential = $credResult.Credential
            $CredRemember = $credResult.Remember
        }
    }

    if ($Credential) {
        $InvokeParams['Credential'] = $Credential
        $InvokeParams['Authentication'] = 'Negotiate'
    }

    # TrustedHosts for IP connections
    if ($IsIP) {
        try {
            $CurrentTrusted = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction Stop).Value
            $TrustedList = if ($CurrentTrusted) { $CurrentTrusted -split ',' | ForEach-Object { $_.Trim() } } else { @() }
            if ($TargetHost -notin $TrustedList -and '*' -notin $TrustedList) {
                if (-not $SkipPrompts) {
                    $AddTrust = [System.Windows.MessageBox]::Show(
                        "IP address '$TargetHost' is not in your WinRM TrustedHosts list.`n`nThis is required for IP-based connections. Add it now?`n`n(Requires running as Administrator)",
                        "TrustedHosts",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Question)
                    if ($AddTrust -ne 'Yes') {
                        Show-Progress $false
                        Set-Status "Connection cancelled — TrustedHosts not updated." "#f9e2af"
                        return $false
                    }
                }
                $NewValue = if ($CurrentTrusted) { "$CurrentTrusted,$TargetHost" } else { $TargetHost }
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value $NewValue -Force -ErrorAction Stop
                Set-Status "Added $TargetHost to TrustedHosts." "#a6e3a1"
                $Window.Dispatcher.Invoke([Action]{}, 'Background')
            }
        }
        catch {
            if (-not $SkipPrompts) {
                [System.Windows.MessageBox]::Show(
                    "Could not check/update TrustedHosts.`n`nError: $($_.Exception.Message)`n`nTry running HyperV Explorer as Administrator, or use a hostname instead of an IP.",
                    "TrustedHosts Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
            Show-Progress $false
            return $false
        }
    }

    # WinRM port check (TCP 5985)
    Set-Status "Checking WinRM port on $TargetHost ..." "#89b4fa"
    $Window.Dispatcher.Invoke([Action]{}, 'Background')

    $PortTest = Test-NetConnection -ComputerName $TargetHost -Port 5985 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    if (-not $PortTest.TcpTestSucceeded) {
        Show-Progress $false
        Set-Status "WinRM port closed on $TargetHost" "#f38ba8"
        if (-not $SkipPrompts) {
            [System.Windows.MessageBox]::Show(
                "WinRM port (TCP 5985) is not open on '$TargetHost'.`n`nThis means WinRM is not enabled or a firewall is blocking it.`n`nOn the target host, run (as Administrator):`n  Enable-PSRemoting -Force",
                "WinRM Port Closed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
        return $false
    }

    # WinRM authentication test
    Set-Status "WinRM port open — testing authentication on $TargetHost ..." "#89b4fa"
    $Window.Dispatcher.Invoke([Action]{}, 'Background')

    $WinRMParams = @{ ComputerName = $TargetHost; ErrorAction = 'Stop' }
    if ($Credential) {
        $WinRMParams['Credential'] = $Credential
        $WinRMParams['Authentication'] = 'Negotiate'
    }
    try {
        Test-WSMan @WinRMParams | Out-Null
    }
    catch {
        Show-Progress $false
        Set-Status "WinRM auth failed on $TargetHost" "#f38ba8"
        if (-not $SkipPrompts) {
            [System.Windows.MessageBox]::Show(
                "WinRM port is open on '$TargetHost' but authentication failed.`n`nError: $($_.Exception.Message)`n`nCheck that:`n- Your credentials are correct`n- Your account has remote management permissions`n- The WinRM service is fully configured on the target",
                "WinRM Authentication Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
        return $false
    }

    Set-Status "WinRM OK — collecting Hyper-V data from $TargetHost ..." "#89b4fa"
    $Window.Dispatcher.Invoke([Action]{}, 'Background')

    # Connect and collect
    try {
        $Results = Invoke-Command @InvokeParams

        if ($null -eq $Results -or @($Results).Count -eq 0) {
            Set-Status "Connected to $TargetHost — no VMs found on this host." "#f9e2af"
            $script:ConnectedHosts[$TargetHost] = @{ Credential = $Credential; VMCount = 0 }
        } else {
            $Count = @($Results).Count
            foreach ($VM in $Results) {
                $script:VMData.Add($VM)
            }
            $script:ConnectedHosts[$TargetHost] = @{ Credential = $Credential; VMCount = $Count }
            Set-Status "Connected to $TargetHost — $Count VMs loaded." "#a6e3a1"
        }

        # Save to history on successful connection
        Save-HostToHistory -Address $TargetHost -UseCurrentUser $UseCurrentUser `
            -Credential $Credential -RememberCredential $CredRemember

        return $true
    }
    catch {
        $ErrMsg = $_.Exception.Message
        Set-Status "Failed to connect to $TargetHost" "#f38ba8"
        if (-not $SkipPrompts) {
            [System.Windows.MessageBox]::Show(
                "Could not collect data from '$TargetHost'.`n`nError: $ErrMsg`n`nMake sure:`n- WinRM is enabled on the target (Enable-PSRemoting -Force)`n- You have Hyper-V admin permissions on the target`n- The host is reachable on the network",
                "Connection Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
        return $false
    }
    finally {
        Show-Progress $false
        Update-StatusBar
    }
}

# ---- Remote collection scriptblock (runs on target host) ----
$CollectionScript = {
    $HostInfo    = Get-VMHost
    $HostName    = $HostInfo.ComputerName
    $HostCPU     = $HostInfo.LogicalProcessorCount
    $HostMemory  = [math]::Round($HostInfo.MemoryCapacity / 1GB, 2)
    $HostVersion = $HostInfo.Version

    $VMs = Get-VM
    foreach ($VM in $VMs) {
        $CPUCount = (Get-VMProcessor -VMName $VM.Name).Count

        $NICs = (Get-VMNetworkAdapter -VMName $VM.Name | ForEach-Object {
            "$($_.Name) [Switch: $($_.SwitchName), MAC: $($_.MacAddress), IPs: $($_.IPAddresses -join ', ')]"
        }) -join "; "

        $Disks = (Get-VMHardDiskDrive -VMName $VM.Name | ForEach-Object {
            try {
                $VHD = Get-VHD -Path $_.Path -ErrorAction Stop
                "$($_.ControllerType)#$($_.ControllerNumber): $($_.Path) (Size: $([math]::Round($VHD.Size / 1GB, 2)) GB, Used: $([math]::Round($VHD.FileSize / 1GB, 2)) GB)"
            } catch {
                "$($_.ControllerType)#$($_.ControllerNumber): $($_.Path) (VHD info unavailable)"
            }
        }) -join "; "

        $Checkpoints = (Get-VMSnapshot -VMName $VM.Name -ErrorAction SilentlyContinue |
                        ForEach-Object { $_.Name }) -join "; "
        if (-not $Checkpoints) { $Checkpoints = "None" }

        $IntegrationServices = $VM.IntegrationServicesVersion

        [PSCustomObject]@{
            HostName            = $HostName
            HostCPU             = $HostCPU
            HostMemoryGB        = $HostMemory
            HostVersion         = "$HostVersion"
            VMName              = $VM.Name
            State               = "$($VM.State)"
            CPUCount            = $CPUCount
            MemoryAssignedMB    = [math]::Round($VM.MemoryAssigned / 1MB, 2)
            Uptime              = "$($VM.Uptime)"
            Generation          = $VM.Generation
            DynamicMemory       = "$($VM.DynamicMemoryEnabled)"
            NICs                = $NICs
            Disks               = $Disks
            Checkpoints         = $Checkpoints
            IntegrationServices = "$IntegrationServices"
        }
    }
}

# ---- Connect button ----
$btnConnect.Add_Click({
    $TargetHost = $txtHost.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($TargetHost)) {
        [System.Windows.MessageBox]::Show("Please enter a hostname or IP address.", "HyperV Explorer",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $btnConnect.IsEnabled = $false
    try {
        $UseCurrentUser = $chkCurrentUser.IsChecked -eq $true
        $result = Connect-HyperVHost -TargetHost $TargetHost -UseCurrentUser $UseCurrentUser
        if ($result) { $txtHost.Text = "" }
    }
    finally {
        $btnConnect.IsEnabled = $true
    }
})

# ---- Bulk Connect button ----
$btnBulkConnect.Add_Click({
    $selectedHosts = Show-BulkConnectDialog
    if (-not $selectedHosts -or $selectedHosts.Count -eq 0) { return }

    $btnConnect.IsEnabled = $false
    $btnBulkConnect.IsEnabled = $false

    $total = $selectedHosts.Count
    $success = 0
    $failed = 0

    for ($i = 0; $i -lt $total; $i++) {
        $host_ = $selectedHosts[$i]
        Set-Status "Bulk connect: $($i + 1) of $total — $host_ ..." "#89b4fa"
        $Window.Dispatcher.Invoke([Action]{}, 'Background')

        # Look up saved settings for this host
        $histEntry = Get-HostHistoryEntry -Address $host_
        $useCurrentUser = if ($histEntry) { [bool]$histEntry.useCurrentUser } else { $true }
        $savedCred = Get-SavedCredential -Address $host_

        $result = Connect-HyperVHost -TargetHost $host_ -UseCurrentUser $useCurrentUser `
            -ProvidedCredential $savedCred -RememberCredential ($null -ne $savedCred) `
            -SkipPrompts $false

        if ($result) { $success++ } else { $failed++ }
    }

    $btnConnect.IsEnabled = $true
    $btnBulkConnect.IsEnabled = $true
    Set-Status "Bulk connect complete: $success connected, $failed failed (of $total)." $(if ($failed -eq 0) { "#a6e3a1" } else { "#f9e2af" })
    Update-StatusBar
})

# ---- History dropdown button ----
$btnHistory.Add_Click({
    Show-HistoryMenu
})

# ---- Disconnect button ----
$btnDisconnect.Add_Click({
    $SelectedItem = $dgVMs.SelectedItem
    if ($null -eq $SelectedItem) {
        [System.Windows.MessageBox]::Show("Select a VM row first to identify which host to disconnect.",
            "HyperV Explorer", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $HostToRemove = $SelectedItem.HostName
    $Confirm = [System.Windows.MessageBox]::Show(
        "Disconnect host '$HostToRemove' and remove all its VMs from the grid?",
        "Confirm Disconnect",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)

    if ($Confirm -eq 'Yes') {
        $ToRemove = @($script:VMData | Where-Object { $_.HostName -eq $HostToRemove })
        foreach ($Item in $ToRemove) {
            $script:VMData.Remove($Item) | Out-Null
        }
        $script:ConnectedHosts.Remove($HostToRemove)
        Set-Status "Disconnected from $HostToRemove." "#f9e2af"
        Update-StatusBar
    }
})

# ---- Export CSV button ----
$btnExport.Add_Click({
    $SaveDialog = [Microsoft.Win32.SaveFileDialog]::new()
    $SaveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $SaveDialog.FileName = "HyperV_Inventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $SaveDialog.Title = "Export Hyper-V Inventory"

    if ($SaveDialog.ShowDialog() -eq $true) {
        try {
            $script:VMData | Export-Csv -Path $SaveDialog.FileName -NoTypeInformation
            Set-Status "Exported to $($SaveDialog.FileName)" "#a6e3a1"
            [System.Windows.MessageBox]::Show("Export complete!`n`n$($SaveDialog.FileName)",
                "Export Successful", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        catch {
            [System.Windows.MessageBox]::Show("Export failed: $($_.Exception.Message)",
                "Export Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# ---- Clear All button ----
$btnClear.Add_Click({
    $Confirm = [System.Windows.MessageBox]::Show(
        "Clear all data and disconnect all hosts?",
        "Confirm Clear",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)

    if ($Confirm -eq 'Yes') {
        $script:VMData.Clear()
        $script:ConnectedHosts.Clear()
        Set-Status "All data cleared." "#f9e2af"
        Update-StatusBar
    }
})

# ---- Allow Enter key to trigger Connect ----
$txtHost.Add_KeyDown({
    if ($_.Key -eq 'Return') {
        $btnConnect.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    }
})

# ---- Show window ----
Update-StatusBar
$Window.ShowDialog() | Out-Null
