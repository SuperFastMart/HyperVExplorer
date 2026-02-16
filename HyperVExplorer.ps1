# =============================================================================================================
# Script:    HyperVExplorer.ps1
# Version:   1.5
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
            $cfg = Get-Content -Path $path -Raw | ConvertFrom-Json
            # Ensure groups property exists (upgrade from v1 config)
            if (-not (Get-Member -InputObject $cfg -Name 'groups' -MemberType NoteProperty)) {
                $cfg | Add-Member -NotePropertyName 'groups' -NotePropertyValue @()
            }
            return $cfg
        } catch {
            return [PSCustomObject]@{ version = 2; hosts = @(); groups = @() }
        }
    }
    return [PSCustomObject]@{ version = 2; hosts = @(); groups = @() }
}

function Save-Config {
    param($Config)
    $path = Get-ConfigPath
    $Config | ConvertTo-Json -Depth 5 | Set-Content -Path $path -Encoding UTF8
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

# ---- Group management functions ----
function Get-GroupForHost {
    param([string]$Address)
    $config = Load-Config
    foreach ($g in @($config.groups)) {
        if ($Address -in @($g.hosts)) {
            return $g
        }
    }
    return $null
}

function Get-GroupCredential {
    param($Group)
    if ($Group -and $Group.username -and $Group.encryptedPassword) {
        try {
            $secPass = $Group.encryptedPassword | ConvertTo-SecureString
            return [System.Management.Automation.PSCredential]::new($Group.username, $secPass)
        } catch {
            return $null
        }
    }
    return $null
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
        <SolidColorBrush x:Key="AccentMauve" Color="#cba6f7"/>

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
                <TextBlock Grid.Column="2" Text="v1.5" FontSize="12"
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
                <Button x:Name="btnGroups" Content="&#x2630; Groups" Margin="8,0,0,0"
                        Background="#3d3654" Foreground="{StaticResource AccentMauve}"
                        ToolTip="Manage host groups and credential mapping"/>
                <Button x:Name="btnDisconnect" Content="&#x2716; Disconnect Selected Host" Margin="8,0,0,0"
                        Background="#4a3644" Foreground="{StaticResource AccentRed}" IsEnabled="False"/>
                <Button x:Name="btnDisconnectAll" Content="&#x2716; Disconnect All" Margin="8,0,0,0"
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
$btnGroups       = $Window.FindName("btnGroups")
$btnDisconnect    = $Window.FindName("btnDisconnect")
$btnDisconnectAll = $Window.FindName("btnDisconnectAll")
$btnExport        = $Window.FindName("btnExport")
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
    $btnExport.IsEnabled        = $hasData
    $btnClear.IsEnabled         = $hasData
    $btnDisconnect.IsEnabled    = $hasData
    $btnDisconnectAll.IsEnabled = $hasData
}

function Set-Status {
    param([string]$Message, [string]$Color = "#a6adc8")
    $txtStatus.Text = $Message
    $txtStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
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

function New-DarkBrush { param([string]$Color) [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color) }

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

    $chkRemember = $null
    if ($ShowRemember) {
        $chkRemember = [System.Windows.Controls.CheckBox]::new()
        $chkRemember.Content = "Remember credentials for this host"
        $chkRemember.IsChecked = $true
        $chkRemember.Foreground = New-DarkBrush "#a6adc8"
        $chkRemember.FontSize = 12
        $rememberPlaceholder.Content = $chkRemember
    }

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

# ---- Group Edit Dialog (New / Edit) ----
function Show-GroupEditDialog {
    param(
        [string]$GroupName = "",
        [string]$Username = "",
        [bool]$UseCurrentUser = $false
    )

    $titleText = if ($GroupName) { "Edit Group" } else { "New Group" }

    [xml]$GrpXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$titleText" Width="450" Height="340"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        Background="#1e1e2e">
    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="$titleText" Foreground="#cdd6f4" FontSize="16"
                   FontWeight="SemiBold" Margin="0,0,0,16"/>
        <Grid Grid.Row="1" Margin="0,0,0,12">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="100"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Group Name:" Foreground="#a6adc8" VerticalAlignment="Center" FontSize="13"/>
            <TextBox x:Name="txtGrpName" Grid.Column="1" FontSize="13"
                     Background="#313244" Foreground="#cdd6f4" BorderBrush="#45475a"
                     Padding="8,6" CaretBrush="#cdd6f4"/>
        </Grid>
        <StackPanel Grid.Row="2" Margin="0,0,0,12">
            <RadioButton x:Name="rdoCurrentUser" Content="Use current user (Kerberos)" GroupName="auth"
                         Foreground="#cdd6f4" FontSize="13" Margin="0,0,0,6"/>
            <RadioButton x:Name="rdoCredentials" Content="Use specific credentials" GroupName="auth"
                         Foreground="#cdd6f4" FontSize="13" IsChecked="True"/>
        </StackPanel>
        <Grid Grid.Row="3" Margin="0,0,0,8" x:Name="credPanel">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="100"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Username:" Foreground="#a6adc8" VerticalAlignment="Center" FontSize="13"/>
            <TextBox x:Name="txtGrpUser" Grid.Column="1" FontSize="13"
                     Background="#313244" Foreground="#cdd6f4" BorderBrush="#45475a"
                     Padding="8,6" CaretBrush="#cdd6f4"/>
        </Grid>
        <Grid Grid.Row="4" Margin="0,0,0,12" x:Name="passPanel">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="100"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Password:" Foreground="#a6adc8" VerticalAlignment="Center" FontSize="13"/>
            <PasswordBox x:Name="txtGrpPass" Grid.Column="1" FontSize="13"
                         Background="#313244" Foreground="#cdd6f4" BorderBrush="#45475a"
                         Padding="8,6" CaretBrush="#cdd6f4"/>
        </Grid>
        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnGrpSave" Content="Save" Width="90" Padding="8,6" Margin="0,0,8,0"
                    Background="#364a63" Foreground="#89b4fa" FontSize="13" IsDefault="True"/>
            <Button x:Name="btnGrpCancel" Content="Cancel" Width="90" Padding="8,6"
                    Background="#45475a" Foreground="#cdd6f4" FontSize="13" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $GrpReader = [System.Xml.XmlNodeReader]::new($GrpXAML)
    $GrpWindow = [Windows.Markup.XamlReader]::Load($GrpReader)
    $GrpWindow.Owner = $Window

    $txtGrpName    = $GrpWindow.FindName("txtGrpName")
    $rdoCurrentUser = $GrpWindow.FindName("rdoCurrentUser")
    $rdoCredentials = $GrpWindow.FindName("rdoCredentials")
    $txtGrpUser    = $GrpWindow.FindName("txtGrpUser")
    $txtGrpPass    = $GrpWindow.FindName("txtGrpPass")
    $credPanel     = $GrpWindow.FindName("credPanel")
    $passPanel     = $GrpWindow.FindName("passPanel")
    $btnGrpSave    = $GrpWindow.FindName("btnGrpSave")
    $btnGrpCancel  = $GrpWindow.FindName("btnGrpCancel")

    # Pre-fill values
    if ($GroupName) { $txtGrpName.Text = $GroupName }
    if ($Username) { $txtGrpUser.Text = $Username }
    if ($UseCurrentUser) {
        $rdoCurrentUser.IsChecked = $true
        $rdoCredentials.IsChecked = $false
    }

    # Toggle credential fields based on radio selection
    $rdoCurrentUser.Add_Checked({
        $credPanel.IsEnabled = $false
        $passPanel.IsEnabled = $false
        $credPanel.Opacity = 0.4
        $passPanel.Opacity = 0.4
    }.GetNewClosure())

    $rdoCredentials.Add_Checked({
        $credPanel.IsEnabled = $true
        $passPanel.IsEnabled = $true
        $credPanel.Opacity = 1.0
        $passPanel.Opacity = 1.0
    }.GetNewClosure())

    # Apply initial state
    if ($UseCurrentUser) {
        $credPanel.IsEnabled = $false
        $passPanel.IsEnabled = $false
        $credPanel.Opacity = 0.4
        $passPanel.Opacity = 0.4
    }

    $btnGrpSave.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtGrpName.Text)) {
            [System.Windows.MessageBox]::Show("Please enter a group name.", "Group",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        if ($rdoCredentials.IsChecked -and [string]::IsNullOrWhiteSpace($txtGrpUser.Text)) {
            [System.Windows.MessageBox]::Show("Please enter a username for the credential set.", "Group",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        $GrpWindow.DialogResult = $true
    }.GetNewClosure())

    $btnGrpCancel.Add_Click({
        $GrpWindow.DialogResult = $false
    }.GetNewClosure())

    $txtGrpName.Focus() | Out-Null
    $dialogResult = $GrpWindow.ShowDialog()

    if ($dialogResult -eq $true) {
        $result = @{
            Name           = $txtGrpName.Text.Trim()
            UseCurrentUser = ($rdoCurrentUser.IsChecked -eq $true)
            Username       = $null
            SecurePassword = $null
        }
        if (-not $result.UseCurrentUser) {
            $result.Username = $txtGrpUser.Text.Trim()
            $result.SecurePassword = $txtGrpPass.SecurePassword
        }
        return $result
    }
    return $null
}

# ---- Manage Groups Dialog ----
function Show-ManageGroupsDialog {
    [xml]$MgrXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Manage Host Groups" Width="650" Height="550"
        WindowStartupLocation="CenterOwner" ResizeMode="CanResize"
        MinWidth="500" MinHeight="400"
        Background="#1e1e2e">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Host Groups" Foreground="#cdd6f4" FontSize="16"
                   FontWeight="SemiBold" Margin="0,0,0,8"/>

        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,6">
            <Button x:Name="btnNewGrp" Content="+ New Group" Padding="12,4" FontSize="12"
                    Background="#3d3654" Foreground="#cba6f7" Margin="0,0,8,0"/>
            <Button x:Name="btnEditGrp" Content="Edit" Padding="12,4" FontSize="12"
                    Background="#45475a" Foreground="#cdd6f4" Margin="0,0,8,0"/>
            <Button x:Name="btnDelGrp" Content="Delete" Padding="12,4" FontSize="12"
                    Background="#4a3644" Foreground="#f38ba8"/>
        </StackPanel>

        <ListBox x:Name="lstGroups" Grid.Row="2" Background="#313244" Foreground="#cdd6f4"
                 BorderBrush="#45475a" FontSize="13" Padding="4"/>

        <TextBlock Grid.Row="3" x:Name="txtGroupDetail" Text="Select a group above"
                   Foreground="#a6adc8" FontSize="13" Margin="0,10,0,4"/>

        <StackPanel Grid.Row="4" Orientation="Horizontal" Margin="0,0,0,6">
            <Button x:Name="btnAddHost" Content="+ Add Host" Padding="12,4" FontSize="12"
                    Background="#364a63" Foreground="#89b4fa" Margin="0,0,8,0"/>
            <Button x:Name="btnRemoveHost" Content="- Remove Host" Padding="12,4" FontSize="12"
                    Background="#45475a" Foreground="#cdd6f4"/>
        </StackPanel>

        <ListBox x:Name="lstGroupHosts" Grid.Row="5" Background="#313244" Foreground="#cdd6f4"
                 BorderBrush="#45475a" FontSize="13" Padding="4"/>

        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnMgrClose" Content="Close" Width="90" Padding="8,6"
                    Background="#45475a" Foreground="#cdd6f4" FontSize="13" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $MgrReader = [System.Xml.XmlNodeReader]::new($MgrXAML)
    $MgrWindow = [Windows.Markup.XamlReader]::Load($MgrReader)
    $MgrWindow.Owner = $Window

    $lstGroups     = $MgrWindow.FindName("lstGroups")
    $lstGroupHosts = $MgrWindow.FindName("lstGroupHosts")
    $txtGroupDetail = $MgrWindow.FindName("txtGroupDetail")
    $btnNewGrp     = $MgrWindow.FindName("btnNewGrp")
    $btnEditGrp    = $MgrWindow.FindName("btnEditGrp")
    $btnDelGrp     = $MgrWindow.FindName("btnDelGrp")
    $btnAddHost    = $MgrWindow.FindName("btnAddHost")
    $btnRemoveHost = $MgrWindow.FindName("btnRemoveHost")
    $btnMgrClose   = $MgrWindow.FindName("btnMgrClose")

    # Helper: refresh group list
    $refreshGroups = {
        $lstGroups.Items.Clear()
        $config = Load-Config
        foreach ($g in @($config.groups)) {
            $hostCount = @($g.hosts).Count
            $authLabel = if ($g.useCurrentUser) { "Current User" }
                         elseif ($g.username) { $g.username }
                         else { "No creds" }
            $item = [System.Windows.Controls.ListBoxItem]::new()
            $item.Content = "$($g.name)  ($hostCount hosts, $authLabel)"
            $item.Tag = $g.name
            $item.Foreground = New-DarkBrush "#cdd6f4"
            $item.FontSize = 13
            $lstGroups.Items.Add($item) | Out-Null
        }
        $lstGroupHosts.Items.Clear()
        $txtGroupDetail.Text = "Select a group above"
    }

    # Helper: refresh hosts for selected group
    $refreshHosts = {
        $lstGroupHosts.Items.Clear()
        $sel = $lstGroups.SelectedItem
        if (-not $sel) {
            $txtGroupDetail.Text = "Select a group above"
            return
        }
        $groupName = $sel.Tag
        $config = Load-Config
        $group = $config.groups | Where-Object { $_.name -eq $groupName } | Select-Object -First 1
        if (-not $group) { return }

        $authLabel = if ($group.useCurrentUser) { "Current User (Kerberos)" }
                     elseif ($group.username) { "Credentials: $($group.username)" }
                     else { "No credentials saved" }
        $txtGroupDetail.Text = "Hosts in '$groupName' — $authLabel"

        foreach ($h in @($group.hosts)) {
            $item = [System.Windows.Controls.ListBoxItem]::new()
            $item.Content = $h
            $item.Tag = $h
            $item.Foreground = New-DarkBrush "#cdd6f4"
            $item.FontSize = 13
            $lstGroupHosts.Items.Add($item) | Out-Null
        }
    }

    # Initialize
    & $refreshGroups

    # Group selection changed
    $lstGroups.Add_SelectionChanged({
        & $refreshHosts
    })

    # New Group
    $btnNewGrp.Add_Click({
        $result = Show-GroupEditDialog
        if ($result) {
            $config = Load-Config
            # Check for duplicate name
            $existing = $config.groups | Where-Object { $_.name -eq $result.Name }
            if ($existing) {
                [System.Windows.MessageBox]::Show("A group named '$($result.Name)' already exists.",
                    "Duplicate Group", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            $newGroup = [PSCustomObject]@{
                name              = $result.Name
                useCurrentUser    = $result.UseCurrentUser
                username          = $result.Username
                encryptedPassword = $null
                hosts             = @()
            }
            if ($result.SecurePassword -and $result.SecurePassword.Length -gt 0) {
                $newGroup.encryptedPassword = ($result.SecurePassword | ConvertFrom-SecureString)
            }
            $config.groups = @($config.groups) + @($newGroup)
            Save-Config $config
            & $refreshGroups
        }
    })

    # Edit Group
    $btnEditGrp.Add_Click({
        $sel = $lstGroups.SelectedItem
        if (-not $sel) {
            [System.Windows.MessageBox]::Show("Select a group first.", "Edit Group",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        $groupName = $sel.Tag
        $config = Load-Config
        $group = $config.groups | Where-Object { $_.name -eq $groupName } | Select-Object -First 1
        if (-not $group) { return }

        $result = Show-GroupEditDialog -GroupName $group.name -Username $group.username `
            -UseCurrentUser ([bool]$group.useCurrentUser)
        if ($result) {
            $group.name = $result.Name
            $group.useCurrentUser = $result.UseCurrentUser
            $group.username = $result.Username
            if ($result.SecurePassword -and $result.SecurePassword.Length -gt 0) {
                $group.encryptedPassword = ($result.SecurePassword | ConvertFrom-SecureString)
            } elseif ($result.UseCurrentUser) {
                $group.encryptedPassword = $null
                $group.username = $null
            }
            Save-Config $config
            & $refreshGroups
        }
    })

    # Delete Group
    $btnDelGrp.Add_Click({
        $sel = $lstGroups.SelectedItem
        if (-not $sel) {
            [System.Windows.MessageBox]::Show("Select a group first.", "Delete Group",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        $groupName = $sel.Tag
        $confirm = [System.Windows.MessageBox]::Show(
            "Delete group '$groupName' and all its host assignments?`n`n(This does not delete the hosts from history.)",
            "Delete Group",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question)
        if ($confirm -eq 'Yes') {
            $config = Load-Config
            $config.groups = @($config.groups | Where-Object { $_.name -ne $groupName })
            Save-Config $config
            & $refreshGroups
        }
    })

    # Add Hosts to Group (bulk - one per line)
    $btnAddHost.Add_Click({
        $sel = $lstGroups.SelectedItem
        if (-not $sel) {
            [System.Windows.MessageBox]::Show("Select a group first.", "Add Hosts",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        $groupName = $sel.Tag

        [xml]$AddXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Add Hosts to $groupName" Width="450" Height="350"
        WindowStartupLocation="CenterOwner" ResizeMode="CanResize"
        MinWidth="350" MinHeight="250"
        Background="#1e1e2e">
    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="Enter hostnames or IPs (one per line):"
                   Foreground="#a6adc8" FontSize="13" Margin="0,0,0,8"/>
        <TextBox x:Name="txtAddHosts" Grid.Row="1" FontSize="13"
                 AcceptsReturn="True" TextWrapping="NoWrap"
                 VerticalScrollBarVisibility="Auto"
                 Background="#313244" Foreground="#cdd6f4" BorderBrush="#45475a"
                 Padding="8,6" CaretBrush="#cdd6f4" Margin="0,0,0,12"/>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnAddOK" Content="Add All" Width="90" Padding="8,6" Margin="0,0,8,0"
                    Background="#364a63" Foreground="#89b4fa" FontSize="13"/>
            <Button x:Name="btnAddCancel" Content="Cancel" Width="80" Padding="8,6"
                    Background="#45475a" Foreground="#cdd6f4" FontSize="13" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
        $AddReader = [System.Xml.XmlNodeReader]::new($AddXAML)
        $AddWindow = [Windows.Markup.XamlReader]::Load($AddReader)
        $AddWindow.Owner = $MgrWindow

        $txtAddHosts = $AddWindow.FindName("txtAddHosts")
        $btnAddOK    = $AddWindow.FindName("btnAddOK")
        $btnAddCancel = $AddWindow.FindName("btnAddCancel")

        $btnAddOK.Add_Click({
            if ([string]::IsNullOrWhiteSpace($txtAddHosts.Text)) {
                [System.Windows.MessageBox]::Show("Please enter at least one hostname or IP.", "Add Hosts",
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            $AddWindow.DialogResult = $true
        }.GetNewClosure())

        $btnAddCancel.Add_Click({
            $AddWindow.DialogResult = $false
        }.GetNewClosure())

        $txtAddHosts.Focus() | Out-Null
        $addResult = $AddWindow.ShowDialog()

        if ($addResult -eq $true) {
            $lines = $txtAddHosts.Text -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
            if ($lines.Count -eq 0) { return }

            $config = Load-Config
            $group = $config.groups | Where-Object { $_.name -eq $groupName } | Select-Object -First 1
            if ($group) {
                $currentHosts = @($group.hosts)
                $added = 0
                $skipped = 0
                foreach ($hostAddr in $lines) {
                    if ($hostAddr -in $currentHosts) {
                        $skipped++
                        continue
                    }
                    # Remove from any other group first
                    foreach ($otherGroup in @($config.groups)) {
                        if ($otherGroup.name -ne $groupName) {
                            $otherGroup.hosts = @($otherGroup.hosts | Where-Object { $_ -ne $hostAddr })
                        }
                    }
                    $currentHosts += $hostAddr
                    $added++
                }
                $group.hosts = $currentHosts
                Save-Config $config
                & $refreshHosts
                & $refreshGroups

                $msg = "$added host(s) added to '$groupName'."
                if ($skipped -gt 0) { $msg += " $skipped already in group (skipped)." }
                [System.Windows.MessageBox]::Show($msg, "Add Hosts",
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        }
    })

    # Remove Host from Group
    $btnRemoveHost.Add_Click({
        $grpSel = $lstGroups.SelectedItem
        $hostSel = $lstGroupHosts.SelectedItem
        if (-not $grpSel -or -not $hostSel) {
            [System.Windows.MessageBox]::Show("Select a group and a host to remove.", "Remove Host",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        $groupName = $grpSel.Tag
        $hostAddr = $hostSel.Tag
        $config = Load-Config
        $group = $config.groups | Where-Object { $_.name -eq $groupName } | Select-Object -First 1
        if ($group) {
            $group.hosts = @($group.hosts | Where-Object { $_ -ne $hostAddr })
            Save-Config $config
            & $refreshHosts
            & $refreshGroups
        }
    })

    $btnMgrClose.Add_Click({
        $MgrWindow.Close()
    })

    $MgrWindow.ShowDialog() | Out-Null
}

# ---- Bulk Connect Dialog (with group support) ----
function Show-BulkConnectDialog {
    $config = Load-Config
    $hostList = @($config.hosts)
    $groupList = @($config.groups)

    if ($hostList.Count -eq 0 -and $groupList.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No saved hosts or groups found.`n`nConnect to hosts individually first to build your history, or create groups via the Groups button.",
            "Bulk Connect",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information)
        return $null
    }

    [xml]$BulkXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Bulk Connect" Width="580" Height="500"
        WindowStartupLocation="CenterOwner" ResizeMode="CanResize"
        MinWidth="450" MinHeight="350"
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

    $checkboxes = @()
    $groupedAddresses = @()

    # Build grouped sections
    foreach ($g in $groupList) {
        $groupHosts = @($g.hosts)
        if ($groupHosts.Count -eq 0) { continue }
        $groupedAddresses += $groupHosts

        $authLabel = if ($g.useCurrentUser) { "Current User" }
                     elseif ($g.username) { $g.username }
                     else { "No creds" }

        # Group header
        $header = [System.Windows.Controls.TextBlock]::new()
        $header.Text = "$($g.name) — $authLabel"
        $header.FontSize = 14
        $header.FontWeight = [System.Windows.FontWeights]::SemiBold
        $header.Foreground = New-DarkBrush "#cba6f7"
        $header.Margin = [System.Windows.Thickness]::new(0, 8, 0, 4)
        $hostListPanel.Children.Add($header) | Out-Null

        foreach ($addr in $groupHosts) {
            $alreadyConnected = $script:ConnectedHosts.ContainsKey($addr)
            $cb = [System.Windows.Controls.CheckBox]::new()
            $cb.IsChecked = (-not $alreadyConnected)
            $cb.IsEnabled = (-not $alreadyConnected)
            $cb.Margin = [System.Windows.Thickness]::new(20, 2, 0, 2)
            $cb.Foreground = New-DarkBrush "#cdd6f4"
            $cb.FontSize = 13
            $cb.Tag = $addr
            $label = $addr
            if ($alreadyConnected) { $label += "  (already connected)" }
            $cb.Content = $label
            $hostListPanel.Children.Add($cb) | Out-Null
            $checkboxes += $cb
        }
    }

    # Ungrouped hosts
    $ungrouped = @($hostList | Where-Object { $_.address -notin $groupedAddresses })
    if ($ungrouped.Count -gt 0) {
        $header = [System.Windows.Controls.TextBlock]::new()
        $header.Text = "Ungrouped Hosts"
        $header.FontSize = 14
        $header.FontWeight = [System.Windows.FontWeights]::SemiBold
        $header.Foreground = New-DarkBrush "#a6adc8"
        $header.Margin = [System.Windows.Thickness]::new(0, 8, 0, 4)
        $hostListPanel.Children.Add($header) | Out-Null

        foreach ($h in $ungrouped) {
            $authInfo = if ($h.useCurrentUser) { "Current User" }
                        elseif ($h.username) { "Saved: $($h.username)" }
                        else { "Will prompt" }
            $lastDate = try { ([datetime]$h.lastConnected).ToString("yyyy-MM-dd HH:mm") } catch { "" }
            $alreadyConnected = $script:ConnectedHosts.ContainsKey($h.address)

            $cb = [System.Windows.Controls.CheckBox]::new()
            $cb.IsChecked = (-not $alreadyConnected)
            $cb.IsEnabled = (-not $alreadyConnected)
            $cb.Margin = [System.Windows.Thickness]::new(20, 2, 0, 2)
            $cb.Foreground = New-DarkBrush "#cdd6f4"
            $cb.FontSize = 13
            $cb.Tag = $h.address
            $label = "$($h.address)  —  $authInfo  —  $lastDate"
            if ($alreadyConnected) { $label += "  (already connected)" }
            $cb.Content = $label
            $hostListPanel.Children.Add($cb) | Out-Null
            $checkboxes += $cb
        }
    }

    if ($checkboxes.Count -eq 0) {
        $emptyMsg = [System.Windows.Controls.TextBlock]::new()
        $emptyMsg.Text = "No hosts available to connect."
        $emptyMsg.Foreground = New-DarkBrush "#a6adc8"
        $emptyMsg.FontSize = 13
        $emptyMsg.Margin = [System.Windows.Thickness]::new(0, 8, 0, 0)
        $hostListPanel.Children.Add($emptyMsg) | Out-Null
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

    # Capture script-level controls as local vars so .GetNewClosure() can see them
    $localTxtHost = $txtHost
    $localChkCurrentUser = $chkCurrentUser

    $menu = [System.Windows.Controls.ContextMenu]::new()
    $menu.Background = New-DarkBrush "#313244"
    $menu.BorderBrush = New-DarkBrush "#45475a"
    $menu.Foreground = New-DarkBrush "#cdd6f4"

    if ($hostList.Count -eq 0) {
        $emptyItem = [System.Windows.Controls.MenuItem]::new()
        $emptyItem.Header = "No recent hosts"
        $emptyItem.IsEnabled = $false
        $emptyItem.Foreground = New-DarkBrush "#a6adc8"
        $menu.Items.Add($emptyItem) | Out-Null
    } else {
        foreach ($h in $hostList) {
            # Check group membership
            $group = Get-GroupForHost -Address $h.address
            $groupTag = if ($group) { " [$($group.name)]" } else { "" }

            $authTag = if ($group -and -not $group.useCurrentUser -and $group.username) { "[Group: $($group.username)]" }
                       elseif ($group -and $group.useCurrentUser) { "[Group: Current User]" }
                       elseif ($h.useCurrentUser) { "[Current User]" }
                       elseif ($h.username) { "[$($h.username)]" }
                       else { "[No saved creds]" }
            $lastDate = try { ([datetime]$h.lastConnected).ToString("MM/dd HH:mm") } catch { "" }

            $item = [System.Windows.Controls.MenuItem]::new()
            $item.Header = "$($h.address)$groupTag  $authTag  $lastDate"
            $item.Tag = $h.address
            $item.Foreground = New-DarkBrush "#cdd6f4"
            $item.Background = New-DarkBrush "#313244"

            $hostAddress = $h.address
            $hostUseCurrentUser = [bool]$h.useCurrentUser
            if ($group) { $hostUseCurrentUser = [bool]$group.useCurrentUser }
            $item.Add_Click({
                $localTxtHost.Text = $hostAddress
                $localChkCurrentUser.IsChecked = $hostUseCurrentUser
            }.GetNewClosure())

            $menu.Items.Add($item) | Out-Null
        }

        $menu.Items.Add([System.Windows.Controls.Separator]::new()) | Out-Null

        # Capture script-level functions as local refs for the non-modal closure
        $fnLoadConfig = ${function:Load-Config}
        $fnSaveConfig = ${function:Save-Config}
        $fnSetStatus  = ${function:Set-Status}

        $clearItem = [System.Windows.Controls.MenuItem]::new()
        $clearItem.Header = "Clear History"
        $clearItem.Foreground = New-DarkBrush "#f38ba8"
        $clearItem.Background = New-DarkBrush "#313244"
        $clearItem.Add_Click({
            $confirm = [System.Windows.MessageBox]::Show(
                "Clear all saved hosts and credentials?`n`n(Groups will not be affected.)",
                "Clear History",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question)
            if ($confirm -eq 'Yes') {
                $cfg = & $fnLoadConfig
                $cfg.hosts = @()
                & $fnSaveConfig $cfg
                & $fnSetStatus "History cleared." "#f9e2af"
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
        # Check group credentials first
        $hostGroup = Get-GroupForHost -Address $TargetHost
        if ($hostGroup) {
            $groupCred = Get-GroupCredential -Group $hostGroup
            if ($groupCred) {
                $Credential = $groupCred
                $CredRemember = $false  # Group creds are already saved at group level
                Set-Status "Using group '$($hostGroup.name)' credentials for $TargetHost ..." "#cba6f7"
                $Window.Dispatcher.Invoke([Action]{}, 'Background')
            }
        }

        # Check per-host saved credentials
        if (-not $Credential) {
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
            }
        }

        # Prompt for new credentials
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

        # Look up group credentials first, then per-host settings
        $hostGroup = Get-GroupForHost -Address $host_
        $savedCred = $null
        $useCurrentUser = $true

        if ($hostGroup) {
            $useCurrentUser = [bool]$hostGroup.useCurrentUser
            if (-not $useCurrentUser) {
                $savedCred = Get-GroupCredential -Group $hostGroup
            }
        } else {
            $histEntry = Get-HostHistoryEntry -Address $host_
            $useCurrentUser = if ($histEntry) { [bool]$histEntry.useCurrentUser } else { $true }
            if (-not $useCurrentUser) {
                $savedCred = Get-SavedCredential -Address $host_
            }
        }

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

# ---- Groups button ----
$btnGroups.Add_Click({
    Show-ManageGroupsDialog
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

# ---- Disconnect All button ----
$btnDisconnectAll.Add_Click({
    $hostCount = $script:ConnectedHosts.Count
    if ($hostCount -eq 0) {
        [System.Windows.MessageBox]::Show("No hosts are currently connected.",
            "HyperV Explorer", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    $Confirm = [System.Windows.MessageBox]::Show(
        "Disconnect all $hostCount host(s) and remove all VMs from the grid?",
        "Confirm Disconnect All",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)

    if ($Confirm -eq 'Yes') {
        $script:VMData.Clear()
        $script:ConnectedHosts.Clear()
        Set-Status "All hosts disconnected." "#f9e2af"
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
