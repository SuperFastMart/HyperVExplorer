# =============================================================================================================
# Script:    HyperVExplorer.ps1
# Version:   1.2
# Purpose:   WPF GUI tool for remote Hyper-V inventory collection across multiple hosts
# Requires:  PowerShell 5.1+, WinRM enabled on target Hyper-V hosts
# =============================================================================================================


Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

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
                <TextBlock Grid.Column="2" Text="v1.2" FontSize="12"
                           Foreground="{StaticResource FgSubtle}" VerticalAlignment="Center"/>
            </Grid>
        </Border>

        <!-- Connection Toolbar -->
        <Border Grid.Row="1" Background="{StaticResource BgDark}" Padding="16,10">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                <TextBlock Text="Host:" Foreground="{StaticResource FgSubtle}" VerticalAlignment="Center"
                           FontSize="13" Margin="0,0,8,0"/>
                <TextBox x:Name="txtHost" Width="250" VerticalAlignment="Center"
                         ToolTip="Enter hostname or IP address"/>
                <CheckBox x:Name="chkCurrentUser" Content="Use current user" IsChecked="True"
                          Margin="16,0,0,0" VerticalAlignment="Center"/>
                <Button x:Name="btnConnect" Content="&#x2B; Connect" Margin="16,0,0,0"
                        Background="#364a63" Foreground="{StaticResource AccentBlue}"/>
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
$btnDisconnect   = $Window.FindName("btnDisconnect")
$btnExport       = $Window.FindName("btnExport")
$btnClear        = $Window.FindName("btnClear")
$dgVMs           = $Window.FindName("dgVMs")
$txtStatus       = $Window.FindName("txtStatus")
$pbProgress      = $Window.FindName("pbProgress")
$txtVMCount      = $Window.FindName("txtVMCount")

# ---- State ----
$script:VMData = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new()
$dgVMs.ItemsSource = $script:VMData
$script:ConnectedHosts = @{}  # HostName -> @{ Credential = $cred; VMCount = N }

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
    param([string]$TargetHost)

    [xml]$CredXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Credentials — $TargetHost" Width="420" Height="240"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        Background="#1e1e2e">
    <Grid Margin="24">
        <Grid.RowDefinitions>
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
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
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

    $script:CredResult = $null

    $credBtnOK.Add_Click({
        $user = $credTxtUser.Text.Trim()
        $pass = $credTxtPass.SecurePassword
        if ([string]::IsNullOrWhiteSpace($user)) {
            [System.Windows.MessageBox]::Show("Please enter a username.", "Credentials",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        $script:CredResult = [System.Management.Automation.PSCredential]::new($user, $pass)
        $CredWindow.DialogResult = $true
        $CredWindow.Close()
    }.GetNewClosure())

    $credBtnCancel.Add_Click({
        $CredWindow.DialogResult = $false
        $CredWindow.Close()
    }.GetNewClosure())

    $credTxtUser.Focus() | Out-Null
    $dialogResult = $CredWindow.ShowDialog()

    if ($dialogResult) { return $script:CredResult }
    return $null
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

    if ($script:ConnectedHosts.ContainsKey($TargetHost)) {
        [System.Windows.MessageBox]::Show("Host '$TargetHost' is already connected.", "HyperV Explorer",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    # ---- Pre-connection checks ----
    Set-Status "Pinging $TargetHost ..." "#89b4fa"
    Show-Progress $true
    $btnConnect.IsEnabled = $false
    $Window.Dispatcher.Invoke([Action]{}, 'Background')

    $PingOK = Test-Connection -ComputerName $TargetHost -Count 2 -Quiet -ErrorAction SilentlyContinue
    if (-not $PingOK) {
        Show-Progress $false
        $btnConnect.IsEnabled = $true
        Set-Status "Host unreachable: $TargetHost" "#f38ba8"
        [System.Windows.MessageBox]::Show(
            "Cannot reach '$TargetHost'.`n`nPing failed. Please check:`n- The hostname or IP is correct`n- The host is powered on and on the network`n- Firewalls allow ICMP",
            "Host Unreachable",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error)
        return
    }
    Set-Status "Ping OK — checking WinRM on $TargetHost ..." "#89b4fa"
    $Window.Dispatcher.Invoke([Action]{}, 'Background')

    # Detect IP address — Kerberos won't work, need credentials + Negotiate
    $IsIP = Test-IsIPAddress $TargetHost
    $UseCurrentUser = $chkCurrentUser.IsChecked

    if ($IsIP -and $UseCurrentUser) {
        $Answer = [System.Windows.MessageBox]::Show(
            "You are connecting by IP address ($($TargetHost)).`n`nKerberos authentication does not work with IP addresses. You need to either:`n`n1. Use a hostname instead of an IP`n2. Provide explicit credentials (click Yes)`n`nWould you like to enter credentials now?",
            "IP Address Detected",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
        if ($Answer -ne 'Yes') { return }
        $UseCurrentUser = $false
    }

    # Build connection parameters
    $InvokeParams = @{
        ComputerName = $TargetHost
        ScriptBlock  = $CollectionScript
        ErrorAction  = 'Stop'
    }

    $Credential = $null
    if (-not $UseCurrentUser) {
        $Credential = Show-CredentialDialog -TargetHost $TargetHost
        if (-not $Credential) {
            Show-Progress $false
            $btnConnect.IsEnabled = $true
            Set-Status "Connection cancelled." "#f9e2af"
            return
        }
        $InvokeParams['Credential'] = $Credential
        $InvokeParams['Authentication'] = 'Negotiate'
    }

    # When connecting by IP, ensure the host is in TrustedHosts
    if ($IsIP) {
        try {
            $CurrentTrusted = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction Stop).Value
            $TrustedList = if ($CurrentTrusted) { $CurrentTrusted -split ',' | ForEach-Object { $_.Trim() } } else { @() }
            if ($TargetHost -notin $TrustedList -and '*' -notin $TrustedList) {
                $AddTrust = [System.Windows.MessageBox]::Show(
                    "IP address '$TargetHost' is not in your WinRM TrustedHosts list.`n`nThis is required for IP-based connections. Add it now?`n`n(Requires running as Administrator)",
                    "TrustedHosts",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Question)
                if ($AddTrust -eq 'Yes') {
                    $NewValue = if ($CurrentTrusted) { "$CurrentTrusted,$TargetHost" } else { $TargetHost }
                    Set-Item WSMan:\localhost\Client\TrustedHosts -Value $NewValue -Force -ErrorAction Stop
                    Set-Status "Added $TargetHost to TrustedHosts." "#a6e3a1"
                    $Window.Dispatcher.Invoke([Action]{}, 'Background')
                } else {
                    Show-Progress $false
                    $btnConnect.IsEnabled = $true
                    Set-Status "Connection cancelled — TrustedHosts not updated." "#f9e2af"
                    return
                }
            }
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Could not check/update TrustedHosts.`n`nError: $($_.Exception.Message)`n`nTry running HyperV Explorer as Administrator, or use a hostname instead of an IP.",
                "TrustedHosts Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error)
            Show-Progress $false
            $btnConnect.IsEnabled = $true
            return
        }
    }

    # WinRM connectivity test
    Set-Status "Testing WinRM on $TargetHost ..." "#89b4fa"
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
        $btnConnect.IsEnabled = $true
        Set-Status "WinRM not available on $TargetHost" "#f38ba8"
        [System.Windows.MessageBox]::Show(
            "WinRM service is not available on '$TargetHost'.`n`nError: $($_.Exception.Message)`n`nOn the target host, run:`n  Enable-PSRemoting -Force`n  winrm quickconfig`n`nAlso check that the Windows Firewall allows WinRM (TCP 5985/5986).",
            "WinRM Unavailable",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error)
        return
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
    }
    catch {
        $ErrMsg = $_.Exception.Message
        Set-Status "Failed to connect to $TargetHost" "#f38ba8"
        [System.Windows.MessageBox]::Show(
            "Could not connect to '$TargetHost'.`n`nError: $ErrMsg`n`nMake sure:`n- WinRM is enabled on the target (Enable-PSRemoting -Force)`n- You have Hyper-V admin permissions on the target`n- The host is reachable on the network`n- If using IP: the app is running as Administrator",
            "Connection Failed",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error)
    }
    finally {
        Show-Progress $false
        $btnConnect.IsEnabled = $true
        Update-StatusBar
        $txtHost.Text = ""
    }
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
