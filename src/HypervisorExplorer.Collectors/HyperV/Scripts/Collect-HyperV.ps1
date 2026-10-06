<#
.SYNOPSIS
    Collects Hyper-V and Failover Cluster inventory for Hypervisor Explorer.

.DESCRIPTION
    Runs ON a Hyper-V host, in Windows PowerShell 5.1 or PowerShell 7, elevated or as a member of
    "Hyper-V Administrators". It is read-only: nothing on the host is changed.

    The output is one JSON document. Hypervisor Explorer runs this script remotely over WinRM, or
    you can run it by hand (for example on an air-gapped host) and import the JSON file in the app.

.PARAMETER OutFile
    Write the JSON to this file (UTF-8, no BOM). Without it the JSON is written to the pipeline.
    Prefer -OutFile over "> file.json": Windows PowerShell 5.1 redirection writes UTF-16.

.PARAMETER SkipCluster
    Do not collect Failover Cluster data (cluster, nodes, networks, CSVs, VM roles).

.PARAMETER ClusterPrimary
    Used by Hypervisor Explorer when it collects every cluster node: cluster-wide data is only
    collected when this computer is the named node, so it is gathered once.

.EXAMPLE
    .\Collect-HyperV.ps1 -OutFile "C:\Temp\$env:COMPUTERNAME-hyperv.json"
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)][string]$ClusterPrimary,
    [string]$OutFile,
    [switch]$SkipCluster
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$SchemaVersion = 1
$ScriptVersion = '1.0.0'
$Warnings = New-Object 'System.Collections.Generic.List[string]'

# Windows PowerShell 5.1: ETS data on System.Array can make ConvertTo-Json emit arrays as {"value":[..],"Count":n}.
Remove-TypeData -TypeName System.Array -ErrorAction SilentlyContinue

# ---------------------------------------------------------------- helpers

function Write-Step([string]$Text) {
    # Progress goes to the verbose stream; Hypervisor Explorer forwards "PROGRESS:" records.
    Write-Verbose -Message ('PROGRESS: ' + $env:COMPUTERNAME + ': ' + $Text) -Verbose
}

function Add-Warning([string]$Text) {
    [void]$Warnings.Add($env:COMPUTERNAME + ': ' + $Text)
}

# String or $null (enums become their names; ConvertTo-Json in 5.1 would emit numbers).
function S($Value) {
    if ($null -eq $Value) { return $null }
    $t = [string]$Value
    if ($t.Length -eq 0) { return $null }
    return $t
}

# Int64 or $null.
function L($Value) {
    if ($null -eq $Value) { return $null }
    try { return [int64]$Value } catch { return $null }
}

function B($Value) {
    if ($null -eq $Value) { return $null }
    return [bool]$Value
}

# ISO-8601 date string or $null.
function D($Value) {
    if ($null -eq $Value) { return $null }
    try { $dt = [datetime]$Value } catch { return $null }
    if ($dt.Year -lt 1900) { return $null }
    return $dt.ToString('o')
}

# Array of non-empty strings (always an array, even with 0 or 1 items).
function Strs($Values) {
    $o = @()
    foreach ($v in @($Values)) {
        if ($null -ne $v) {
            $t = [string]$v
            if ($t.Length -gt 0) { $o += $t }
        }
    }
    return , $o
}

# Name of a cluster object (node, group, resource type) whether it arrives as an object or a string.
function N($Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }
    $n = $null
    try { $n = $Value.Name } catch { $n = $null }
    if ($n) { return [string]$n }
    return S $Value
}

function Get-Cim([string]$Class, [string]$Namespace = 'root\cimv2', [string]$Filter) {
    $p = @{ ClassName = $Class; Namespace = $Namespace; ErrorAction = 'Stop' }
    if ($Filter) { $p['Filter'] = $Filter }
    if (Get-Command -Name Get-CimInstance -ErrorAction SilentlyContinue) { return @(Get-CimInstance @p) }
    $w = @{ Class = $Class; Namespace = $Namespace; ErrorAction = 'Stop' }
    if ($Filter) { $w['Filter'] = $Filter }
    return @(Get-WmiObject @w)
}

# ---------------------------------------------------------------- pre-flight

if (-not (Get-Command -Name Get-VM -Module Hyper-V -ErrorAction SilentlyContinue)) {
    throw ('HYPERV_MODULE_MISSING: The Hyper-V PowerShell module is not available on ' + $env:COMPUTERNAME +
        '. Is this a Hyper-V host? Install it with: Install-WindowsFeature Hyper-V-PowerShell (Windows Server) or ' +
        'Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-Management-PowerShell (Windows client).')
}

Write-Step 'Reading host configuration'
try {
    $vmh = Hyper-V\Get-VMHost
} catch {
    throw ('HYPERV_ACCESS: Get-VMHost failed on ' + $env:COMPUTERNAME + ': ' + $_.Exception.Message)
}

# ---------------------------------------------------------------- host

$cs = $null; $os = $null; $bios = $null; $csp = $null; $procs = @()
try { $cs = @(Get-Cim 'Win32_ComputerSystem')[0] } catch { Add-Warning ('Win32_ComputerSystem: ' + $_.Exception.Message) }
try { $os = @(Get-Cim 'Win32_OperatingSystem')[0] } catch { Add-Warning ('Win32_OperatingSystem: ' + $_.Exception.Message) }
try { $bios = @(Get-Cim 'Win32_BIOS')[0] } catch { Add-Warning ('Win32_BIOS: ' + $_.Exception.Message) }
try { $csp = @(Get-Cim 'Win32_ComputerSystemProduct')[0] } catch { Add-Warning ('Win32_ComputerSystemProduct: ' + $_.Exception.Message) }
try { $procs = @(Get-Cim 'Win32_Processor') } catch { Add-Warning ('Win32_Processor: ' + $_.Exception.Message) }

$hostName = S $cs.Name
if (-not $hostName) { $hostName = $env:COMPUTERNAME }

$processors = @()
foreach ($p in $procs) {
    $processors += [ordered]@{
        socket            = S $p.SocketDesignation
        name              = S $p.Name
        manufacturer      = S $p.Manufacturer
        maxClockSpeed     = L $p.MaxClockSpeed
        numberOfCores     = L $p.NumberOfCores
        numberOfLogicalProcessors = L $p.NumberOfLogicalProcessors
        loadPercentage    = L $p.LoadPercentage
    }
}

$timeZone = $null
try {
    $tz = Get-TimeZone -ErrorAction Stop
    $timeZone = [ordered]@{ id = S $tz.Id; displayName = S $tz.DisplayName; baseUtcOffsetMinutes = L $tz.BaseUtcOffset.TotalMinutes }
} catch {
    try {
        $wtz = @(Get-Cim 'Win32_TimeZone')[0]
        $timeZone = [ordered]@{ id = S $wtz.StandardName; displayName = S $wtz.Caption; baseUtcOffsetMinutes = L $wtz.Bias }
    } catch { Add-Warning ('Time zone: ' + $_.Exception.Message) }
}

$dnsServers = @()
try {
    foreach ($d in @(Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction Stop)) {
        foreach ($s in @($d.ServerAddresses)) {
            if ($s -and ($dnsServers -notcontains [string]$s)) { $dnsServers += [string]$s }
        }
    }
} catch { Add-Warning ('DNS servers: ' + $_.Exception.Message) }

$dnsSearch = @()
try { $dnsSearch = Strs (Get-DnsClientGlobalSetting -ErrorAction Stop).SuffixSearchList } catch { }

$ntpSource = $null
try {
    $src = & w32tm.exe /query /source 2>$null
    if ($LASTEXITCODE -eq 0 -and $src) { $ntpSource = ((@($src) -join ' ').Trim()) }
} catch { }

# Physical NICs
Write-Step 'Reading network configuration'
$pci = @{}
try {
    foreach ($h in @(Get-NetAdapterHardwareInfo -ErrorAction Stop)) {
        try { $pci[[string]$h.Name] = ('{0:x4}:{1:x2}:{2:x2}.{3:x}' -f [int]$h.Segment, [int]$h.Bus, [int]$h.Device, [int]$h.Function) } catch { }
    }
} catch { }

$nics = @()
try {
    foreach ($a in @(Get-NetAdapter -Physical -ErrorAction Stop)) {
        $nics += [ordered]@{
            name                 = S $a.Name
            interfaceDescription = S $a.InterfaceDescription
            driverProvider       = S $a.DriverProvider
            driverDescription    = S $a.DriverDescription
            driverFileName       = S $a.DriverFileName
            driverVersion        = S $a.DriverVersionString
            linkSpeed            = S $a.LinkSpeed
            speedBps             = L $a.Speed
            fullDuplex           = B $a.FullDuplex
            macAddress           = S $a.MacAddress
            status               = S $a.Status
            mtu                  = L $a.MtuSize
            ifIndex              = L $a.ifIndex
            pci                  = $pci[[string]$a.Name]
        }
    }
} catch { Add-Warning ('Physical NICs: ' + $_.Exception.Message) }

$lbfoTeams = @()
try {
    if (Get-Command -Name Get-NetLbfoTeam -ErrorAction SilentlyContinue) {
        foreach ($t in @(Get-NetLbfoTeam -ErrorAction Stop)) {
            $tnicDesc = @()
            try { $tnicDesc = Strs (@(Get-NetAdapter -Name @($t.TeamNics) -ErrorAction Stop) | ForEach-Object { $_.InterfaceDescription }) } catch { }
            $lbfoTeams += [ordered]@{
                name                   = S $t.Name
                members                = Strs $t.Members
                teamingMode            = S $t.TeamingMode
                loadBalancingAlgorithm = S $t.LoadBalancingAlgorithm
                teamNicDescriptions    = $tnicDesc
            }
        }
    }
} catch { }

$switches = @()
try {
    foreach ($sw in @(Get-VMSwitch -ErrorAction Stop)) {
        $descs = @()
        if ($sw.NetAdapterInterfaceDescriptions) { $descs = Strs $sw.NetAdapterInterfaceDescriptions }
        elseif ($sw.NetAdapterInterfaceDescription) { $descs = Strs $sw.NetAdapterInterfaceDescription }
        $switches += [ordered]@{
            name                            = S $sw.Name
            id                              = S $sw.Id
            switchType                      = S $sw.SwitchType
            netAdapterInterfaceDescriptions = $descs
            embeddedTeamingEnabled          = B $sw.EmbeddedTeamingEnabled
            allowManagementOS               = B $sw.AllowManagementOS
            bandwidthReservationMode        = S $sw.BandwidthReservationMode
            notes                           = S $sw.Notes
        }
    }
} catch { Add-Warning ('Virtual switches: ' + $_.Exception.Message) }

$mgmtAdapters = @()
try {
    foreach ($m in @(Get-VMNetworkAdapter -ManagementOS -ErrorAction Stop)) {
        $mode = $null; $vlan = $null
        try {
            $v = Get-VMNetworkAdapterVlan -VMNetworkAdapter $m -ErrorAction Stop
            $mode = S $v.OperationMode
            $vlan = L $v.AccessVlanId
        } catch { }
        $mgmtAdapters += [ordered]@{
            name         = S $m.Name
            switchName   = S $m.SwitchName
            macAddress   = S $m.MacAddress
            status       = S (@($m.Status) -join ',')
            vlanMode     = $mode
            accessVlanId = $vlan
        }
    }
} catch { Add-Warning ('Management OS adapters: ' + $_.Exception.Message) }

$ipInterfaces = @()
try {
    $adapters = @{}
    foreach ($a in @(Get-NetAdapter -ErrorAction SilentlyContinue)) { $adapters[[int]$a.ifIndex] = $a }
    $ipIf = @{}
    foreach ($i in @(Get-NetIPInterface -AddressFamily IPv4 -ErrorAction SilentlyContinue)) { $ipIf[[int]$i.InterfaceIndex] = $i }
    $gw = @{}
    foreach ($r in @(Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue)) {
        if (-not $gw.ContainsKey([int]$r.InterfaceIndex)) { $gw[[int]$r.InterfaceIndex] = [string]$r.NextHop }
    }
    $addrs = @(Get-NetIPAddress -ErrorAction Stop | Where-Object { [string]$_.InterfaceAlias -notlike 'Loopback*' })
    foreach ($grp in @($addrs | Group-Object -Property InterfaceIndex)) {
        $idx = [int]$grp.Name
        $v4 = @(); $v6 = @()
        foreach ($ad in $grp.Group) {
            $ip = ([string]$ad.IPAddress).Split('%')[0]
            if ([string]$ad.AddressFamily -eq 'IPv4') {
                if ($ip -notlike '169.254.*') { $v4 += [ordered]@{ address = $ip; prefixLength = L $ad.PrefixLength; origin = S $ad.PrefixOrigin } }
            } elseif ($ip -notlike 'fe80*') {
                $v6 += [ordered]@{ address = $ip; prefixLength = L $ad.PrefixLength; origin = S $ad.PrefixOrigin }
            }
        }
        if ($v4.Count -eq 0 -and $v6.Count -eq 0) { continue }
        $a = $adapters[$idx]
        $ii = $ipIf[$idx]
        $dhcp = $null; $mtu = $null
        if ($ii) { $dhcp = ([string]$ii.Dhcp -eq 'Enabled'); $mtu = L $ii.NlMtu }
        $ipInterfaces += [ordered]@{
            interfaceAlias       = S (@($grp.Group)[0].InterfaceAlias)
            ifIndex              = $idx
            interfaceDescription = S $a.InterfaceDescription
            macAddress           = S $a.MacAddress
            dhcp                 = $dhcp
            mtu                  = $mtu
            gateway              = $gw[$idx]
            ipv4                 = $v4
            ipv6                 = $v6
        }
    }
} catch { Add-Warning ('IP configuration: ' + $_.Exception.Message) }

# Storage adapters
$initiatorPorts = @()
try {
    if (Get-Command -Name Get-InitiatorPort -ErrorAction SilentlyContinue) {
        foreach ($ip in @(Get-InitiatorPort -ErrorAction Stop)) {
            $initiatorPorts += [ordered]@{
                instanceName      = S $ip.InstanceName
                nodeAddress       = S $ip.NodeAddress
                portAddress       = S $ip.PortAddress
                connectionType    = S $ip.ConnectionType
                operationalStatus = S (@($ip.OperationalStatus) -join ',')
            }
        }
    }
} catch { }

$scsiControllers = @()
try {
    foreach ($c in @(Get-Cim 'Win32_SCSIController')) {
        $scsiControllers += [ordered]@{
            name         = S $c.Name
            driverName   = S $c.DriverName
            manufacturer = S $c.Manufacturer
            status       = S $c.Status
            pnpDeviceId  = S $c.PNPDeviceID
        }
    }
} catch { }

# Local volumes (become datastores). CSV volumes are reported by the cluster section.
Write-Step 'Reading volumes'
$volumes = @()
try {
    foreach ($v in @(Get-Volume -ErrorAction Stop)) {
        if ([string]$v.DriveType -ne 'Fixed') { continue }
        if ([string]$v.FileSystem -like 'CSVFS*') { continue }
        if (-not $v.Size) { continue }
        $paths = @()
        if ($v.DriveLetter) { $paths += ([string]$v.DriveLetter + ':\') }
        try {
            foreach ($ap in @(($v | Get-Partition -ErrorAction Stop).AccessPaths)) {
                if ($ap -and -not ([string]$ap).StartsWith('\\?\') -and ($paths -notcontains [string]$ap)) { $paths += [string]$ap }
            }
        } catch { }
        if ($paths.Count -eq 0) { continue }
        $volumes += [ordered]@{
            label         = S $v.FileSystemLabel
            fileSystem    = S $v.FileSystem
            driveLetter   = S $v.DriveLetter
            size          = L $v.Size
            sizeRemaining = L $v.SizeRemaining
            healthStatus  = S $v.HealthStatus
            paths         = $paths
        }
    }
} catch { Add-Warning ('Volumes: ' + $_.Exception.Message) }

$hostData = [ordered]@{
    name                      = $hostName
    dnsHostName               = S $cs.DNSHostName
    domain                    = S $cs.Domain
    partOfDomain              = B $cs.PartOfDomain
    manufacturer              = S $cs.Manufacturer
    model                     = S $cs.Model
    numberOfProcessors        = L $cs.NumberOfProcessors
    totalPhysicalMemory       = L $cs.TotalPhysicalMemory
    logicalProcessorCount     = L $vmh.LogicalProcessorCount
    memoryCapacity            = L $vmh.MemoryCapacity
    virtualHardDiskPath       = S $vmh.VirtualHardDiskPath
    virtualMachinePath        = S $vmh.VirtualMachinePath
    numaSpanningEnabled       = B $vmh.NumaSpanningEnabled
    enableEnhancedSessionMode = B $vmh.EnableEnhancedSessionMode
    liveMigrationEnabled      = B $vmh.VirtualMachineMigrationEnabled
    maximumLiveMigrations     = L $vmh.MaximumVirtualMachineMigrations
    processors                = $processors
    os                        = [ordered]@{
        caption                = S $os.Caption
        version                = S $os.Version
        buildNumber            = S $os.BuildNumber
        lastBootUpTime         = D $os.LastBootUpTime
        installDate            = D $os.InstallDate
        freePhysicalMemoryKb   = L $os.FreePhysicalMemory
        totalVisibleMemoryKb   = L $os.TotalVisibleMemorySize
    }
    bios                      = [ordered]@{
        manufacturer = S $bios.Manufacturer
        version      = S $bios.SMBIOSBIOSVersion
        releaseDate  = D $bios.ReleaseDate
        serialNumber = S $bios.SerialNumber
    }
    uuid                      = S $csp.UUID
    timeZone                  = $timeZone
    dnsServers                = $dnsServers
    dnsSearch                 = $dnsSearch
    ntpSource                 = $ntpSource
    nics                      = $nics
    lbfoTeams                 = $lbfoTeams
    switches                  = $switches
    managementAdapters        = $mgmtAdapters
    ipInterfaces              = $ipInterfaces
    initiatorPorts            = $initiatorPorts
    scsiControllers           = $scsiControllers
    volumes                   = $volumes
}

# ---------------------------------------------------------------- failover cluster

$clusterData = $null
$collectCluster = -not $SkipCluster
if ($ClusterPrimary -and -not [string]::Equals($ClusterPrimary, $env:COMPUTERNAME, [StringComparison]::OrdinalIgnoreCase)) {
    $collectCluster = $false
}
if ($collectCluster -and (Get-Command -Name Get-Cluster -ErrorAction SilentlyContinue)) {
    $cl = $null
    try { $cl = Get-Cluster -ErrorAction Stop } catch {
        $svc = Get-Service -Name ClusSvc -ErrorAction SilentlyContinue
        if ($svc) { Add-Warning ('Failover cluster: ' + $_.Exception.Message) }
    }
    if ($cl) {
        Write-Step ('Reading failover cluster ' + $cl.Name)
        $quorumType = $null; $witness = $null; $witnessType = $null
        try {
            $q = Get-ClusterQuorum -ErrorAction Stop
            $quorumType = S $q.QuorumType
            if ($q.QuorumResource) {
                $witness = N $q.QuorumResource
                $witnessType = N $q.QuorumResource.ResourceType
            }
        } catch { Add-Warning ('Cluster quorum: ' + $_.Exception.Message) }

        $cNodes = @()
        try {
            foreach ($n in @(Get-ClusterNode -ErrorAction Stop)) {
                $cNodes += [ordered]@{
                    name          = S $n.Name
                    id            = S $n.Id
                    state         = S $n.State
                    drainStatus   = S $n.DrainStatus
                    nodeWeight    = L $n.NodeWeight
                    dynamicWeight = L $n.DynamicWeight
                    manufacturer  = S $n.Manufacturer
                    model         = S $n.Model
                    serialNumber  = S $n.SerialNumber
                }
            }
        } catch { Add-Warning ('Cluster nodes: ' + $_.Exception.Message) }

        $cNetIfs = @()
        try {
            foreach ($ni in @(Get-ClusterNetworkInterface -ErrorAction Stop)) {
                $cNetIfs += [ordered]@{
                    name    = S $ni.Name
                    node    = N $ni.Node
                    network = N $ni.Network
                    address = S $ni.Address
                    state   = S $ni.State
                }
            }
        } catch { }

        $cNets = @()
        try {
            foreach ($cn in @(Get-ClusterNetwork -ErrorAction Stop)) {
                $roleValue = $null
                try { $roleValue = [int64][int]$cn.Role } catch { }
                $cNets += [ordered]@{
                    name              = S $cn.Name
                    id                = S $cn.Id
                    role              = S $cn.Role
                    roleValue         = $roleValue
                    address           = S $cn.Address
                    addressMask       = S $cn.AddressMask
                    ipv4PrefixLengths = Strs $cn.Ipv4PrefixLengths
                    state             = S $cn.State
                    metric            = L $cn.Metric
                    autoMetric        = B $cn.AutoMetric
                }
            }
        } catch { Add-Warning ('Cluster networks: ' + $_.Exception.Message) }

        $csvs = @()
        try {
            foreach ($csv in @(Get-ClusterSharedVolume -ErrorAction Stop)) {
                foreach ($svi in @($csv.SharedVolumeInfo)) {
                    if ($null -eq $svi) { continue }
                    $part = $svi.Partition
                    $csvs += [ordered]@{
                        name               = S $csv.Name
                        id                 = S $csv.Id
                        state              = S $csv.State
                        ownerNode          = N $csv.OwnerNode
                        friendlyVolumeName = S $svi.FriendlyVolumeName
                        fileSystem         = S $part.FileSystem
                        size               = L $part.Size
                        freeSpace          = L $part.FreeSpace
                        maintenanceMode    = B $svi.MaintenanceMode
                        redirectedAccess   = B $svi.RedirectedAccess
                        faultState         = S $svi.FaultState
                    }
                }
            }
        } catch { Add-Warning ('Cluster Shared Volumes: ' + $_.Exception.Message) }

        # VM resource -> VmId, so roles can be matched to VMs even when the role name differs.
        $vmIdByGroup = @{}
        try {
            $vmRes = @(Get-ClusterResource -ErrorAction Stop | Where-Object { (N $_.ResourceType) -eq 'Virtual Machine' })
            foreach ($prm in @($vmRes | Get-ClusterParameter -Name VmId -ErrorAction SilentlyContinue)) {
                $grpName = N $prm.ClusterObject.OwnerGroup
                if ($grpName -and $prm.Value) { $vmIdByGroup[$grpName] = ([string]$prm.Value).ToLowerInvariant() }
            }
        } catch { Add-Warning ('Cluster VM resources: ' + $_.Exception.Message) }

        $groups = @()
        try {
            foreach ($g in @(Get-ClusterGroup -ErrorAction Stop)) {
                if ((S $g.GroupType) -ne 'VirtualMachine') { continue }
                $owners = @()
                try { $owners = Strs (@((Get-ClusterOwnerNode -Group ([string]$g.Name) -ErrorAction Stop).OwnerNodes) | ForEach-Object { N $_ }) } catch { }
                $groups += [ordered]@{
                    name              = S $g.Name
                    id                = S $g.Id
                    ownerNode         = N $g.OwnerNode
                    state             = S $g.State
                    priority          = L $g.Priority
                    preferredOwners   = $owners
                    autoFailbackType  = L $g.AutoFailbackType
                    failoverThreshold = L $g.FailoverThreshold
                    failoverPeriod    = L $g.FailoverPeriod
                    vmId              = $vmIdByGroup[[string]$g.Name]
                }
            }
        } catch { Add-Warning ('Cluster roles: ' + $_.Exception.Message) }

        $clusterData = [ordered]@{
            name              = S $cl.Name
            id                = S $cl.Id
            domain            = S $cl.Domain
            functionalLevel   = L $cl.ClusterFunctionalLevel
            s2dEnabled        = L $cl.S2DEnabled
            quorumType        = $quorumType
            quorumWitness     = $witness
            quorumWitnessType = $witnessType
            nodes             = $cNodes
            networkInterfaces = $cNetIfs
            networks          = $cNets
            sharedVolumes     = $csvs
            vmGroups          = $groups
        }
    }
}

# ---------------------------------------------------------------- virtual machines

Write-Step 'Enumerating virtual machines'
try {
    $allVms = @(Hyper-V\Get-VM -ErrorAction Stop)
} catch {
    throw ('HYPERV_ACCESS: Get-VM failed on ' + $env:COMPUTERNAME + ': ' + $_.Exception.Message)
}

# Guest KVP data (OS name, FQDN) for all VMs in one query.
$kvp = @{}
try {
    foreach ($k in @(Get-Cim -Class 'Msvm_KvpExchangeComponent' -Namespace 'root\virtualization\v2')) {
        $items = @{}
        foreach ($x in @($k.GuestIntrinsicExchangeItems)) {
            if (-not $x) { continue }
            try {
                $doc = [xml]$x
                $nm = $null; $val = $null
                foreach ($prop in @($doc.INSTANCE.PROPERTY)) {
                    if ($prop.NAME -eq 'Name') { $nm = [string]$prop.VALUE }
                    elseif ($prop.NAME -eq 'Data') { $val = [string]$prop.VALUE }
                }
                if ($nm) { $items[$nm] = $val }
            } catch { }
        }
        $kvp[([string]$k.SystemName).ToLowerInvariant()] = $items
    }
} catch { Add-Warning ('Guest KVP data unavailable: ' + $_.Exception.Message) }

# BIOS GUID (SMBIOS UUID seen by the guest) for all VMs in one query.
$biosGuid = @{}
try {
    foreach ($s in @(Get-Cim -Class 'Msvm_VirtualSystemSettingData' -Namespace 'root\virtualization\v2' -Filter "VirtualSystemType = 'Microsoft:Hyper-V:System:Realized'")) {
        if ($s.BIOSGUID) {
            $biosGuid[([string]$s.VirtualSystemIdentifier).ToLowerInvariant()] = ([string]$s.BIOSGUID).Trim('{', '}').ToLowerInvariant()
        }
    }
} catch { Add-Warning ('VM BIOS GUIDs unavailable: ' + $_.Exception.Message) }

$heartbeatId = '84EAAE65-2F2E-45F5-9BB5-0E857DC8EB47'

function Get-VmRecord($vm) {
    $id = ([string]$vm.Id).ToLowerInvariant()
    $rec = [ordered]@{
        name                       = S $vm.Name
        id                         = $id
        state                      = S $vm.State
        status                     = S $vm.Status
        uptimeSeconds              = L $vm.Uptime.TotalSeconds
        generation                 = L $vm.Generation
        version                    = S $vm.Version
        processorCount             = L $vm.ProcessorCount
        memoryStartup              = L $vm.MemoryStartup
        memoryAssigned             = L $vm.MemoryAssigned
        memoryDemand               = L $vm.MemoryDemand
        memoryMinimum              = L $vm.MemoryMinimum
        memoryMaximum              = L $vm.MemoryMaximum
        dynamicMemoryEnabled       = B $vm.DynamicMemoryEnabled
        path                       = S $vm.Path
        configurationLocation      = S $vm.ConfigurationLocation
        snapshotFileLocation       = S $vm.SnapshotFileLocation
        smartPagingFilePath        = S $vm.SmartPagingFilePath
        notes                      = S $vm.Notes
        creationTime               = D $vm.CreationTime
        automaticStartAction       = S $vm.AutomaticStartAction
        automaticStartDelay        = L $vm.AutomaticStartDelay
        automaticStopAction        = S $vm.AutomaticStopAction
        isClustered                = B $vm.IsClustered
        replicationState           = S $vm.ReplicationState
        replicationHealth          = S $vm.ReplicationHealth
        checkpointType             = S $vm.CheckpointType
        integrationServicesVersion = S $vm.IntegrationServicesVersion
        integrationServicesState   = S $vm.IntegrationServicesState
        heartbeat                  = S $vm.Heartbeat
        heartbeatStatus            = $null
        integrationServices        = @()
        processor                  = $null
        firmware                   = $null
        bios                       = $null
        disks                      = @()
        dvdDrives                  = @()
        networkAdapters            = @()
        snapshots                  = @()
        guest                      = $null
        biosGuid                   = $biosGuid[$id]
    }

    # Integration services
    $ics = @()
    try {
        foreach ($s in @(Get-VMIntegrationService -VM $vm -ErrorAction Stop)) {
            if ($null -eq $s) { continue }
            $ic = [ordered]@{
                name                       = S $s.Name
                enabled                    = B $s.Enabled
                primaryOperationalStatus   = S $s.PrimaryOperationalStatus
                primaryStatusDescription   = S $s.PrimaryStatusDescription
                secondaryStatusDescription = S $s.SecondaryStatusDescription
            }
            $ics += $ic
            if ($ic.name -eq 'Heartbeat' -or ([string]$s.Id) -like ('*' + $heartbeatId + '*')) {
                $rec.heartbeatStatus = $ic.primaryStatusDescription
                if ($ic.primaryOperationalStatus) { $rec.heartbeatStatus = $ic.primaryOperationalStatus }
            }
        }
    } catch { Add-Warning ("VM '" + $vm.Name + "' integration services: " + $_.Exception.Message) }
    $rec.integrationServices = $ics

    # Processor
    try {
        $p = Get-VMProcessor -VM $vm -ErrorAction Stop
        $rec.processor = [ordered]@{
            count                            = L $p.Count
            reserve                          = L $p.Reserve
            maximum                          = L $p.Maximum
            relativeWeight                   = L $p.RelativeWeight
            exposeVirtualizationExtensions   = B $p.ExposeVirtualizationExtensions
            compatibilityForMigrationEnabled = B $p.CompatibilityForMigrationEnabled
        }
    } catch { Add-Warning ("VM '" + $vm.Name + "' processor: " + $_.Exception.Message) }

    # Firmware (gen 2) / BIOS (gen 1)
    if ($vm.Generation -ge 2) {
        try {
            $fw = Get-VMFirmware -VM $vm -ErrorAction Stop
            $boot = @()
            foreach ($bo in @($fw.BootOrder)) {
                if ($null -eq $bo) { continue }
                $t = [string]$bo.BootType
                if ($t -eq 'File') { $t = 'File:' + [string]$bo.Description }
                elseif ($bo.Device) { $t = $bo.Device.GetType().Name }
                $boot += $t
            }
            $rec.firmware = [ordered]@{
                secureBoot         = ([string]$fw.SecureBoot -eq 'On')
                secureBootTemplate = S $fw.SecureBootTemplate
                bootOrder          = $boot
            }
        } catch { Add-Warning ("VM '" + $vm.Name + "' firmware: " + $_.Exception.Message) }
    } else {
        try {
            $vb = Get-VMBios -VM $vm -ErrorAction Stop
            $rec.bios = [ordered]@{ startupOrder = Strs $vb.StartupOrder }
        } catch { }
    }

    # Hard disks
    $hdds = $vm.HardDrives
    if ($null -eq $hdds) { $hdds = @(Get-VMHardDiskDrive -VM $vm -ErrorAction SilentlyContinue) }
    $disks = @()
    foreach ($d in @($hdds)) {
        if ($null -eq $d) { continue }
        $dr = [ordered]@{
            name                          = S $d.Name
            controllerType                = S $d.ControllerType
            controllerNumber              = L $d.ControllerNumber
            controllerLocation            = L $d.ControllerLocation
            path                          = S $d.Path
            diskNumber                    = L $d.DiskNumber
            supportPersistentReservations = B $d.SupportPersistentReservations
            vhdFormat                     = $null
            vhdType                       = $null
            size                          = $null
            fileSize                      = $null
            parentPath                    = $null
        }
        if ($null -eq $d.DiskNumber -and $d.Path) {
            try {
                $vhd = Get-VHD -Path $d.Path -ErrorAction Stop
                $dr.vhdFormat = S $vhd.VhdFormat
                $dr.vhdType = S $vhd.VhdType
                $dr.size = L $vhd.Size
                $dr.fileSize = L $vhd.FileSize
                $dr.parentPath = S $vhd.ParentPath
            } catch { Add-Warning ("VM '" + $vm.Name + "' disk '" + $d.Path + "': " + $_.Exception.Message) }
        }
        $disks += $dr
    }
    $rec.disks = $disks

    # DVD drives
    $dvds = $vm.DVDDrives
    if ($null -eq $dvds) { $dvds = @(Get-VMDvdDrive -VM $vm -ErrorAction SilentlyContinue) }
    $dv = @()
    foreach ($c in @($dvds)) {
        if ($null -eq $c) { continue }
        $dv += [ordered]@{
            controllerType     = S $c.ControllerType
            controllerNumber   = L $c.ControllerNumber
            controllerLocation = L $c.ControllerLocation
            path               = S $c.Path
            dvdMediaType       = S $c.DvdMediaType
        }
    }
    $rec.dvdDrives = $dv

    # Network adapters
    $vnics = $vm.NetworkAdapters
    if ($null -eq $vnics) { $vnics = @(Get-VMNetworkAdapter -VM $vm -ErrorAction SilentlyContinue) }
    $na = @()
    foreach ($n in @($vnics)) {
        if ($null -eq $n) { continue }
        $vs = $n.VlanSetting
        if ($null -eq $vs) { try { $vs = Get-VMNetworkAdapterVlan -VMNetworkAdapter $n -ErrorAction Stop } catch { $vs = $null } }
        $mode = $null; $vlan = $null; $native = $null; $allowed = $null
        if ($vs) {
            $mode = S $vs.OperationMode
            $vlan = L $vs.AccessVlanId
            $native = L $vs.NativeVlanId
            $allowed = S $vs.AllowedVlanIdListString
        }
        $na += [ordered]@{
            name                     = S $n.Name
            id                       = S $n.Id
            switchName               = S $n.SwitchName
            macAddress               = S $n.MacAddress
            ipAddresses              = Strs $n.IPAddresses
            status                   = S (@($n.Status) -join ',')
            connected                = B $n.Connected
            dynamicMacAddressEnabled = B $n.DynamicMacAddressEnabled
            isLegacy                 = B $n.IsLegacy
            vlanMode                 = $mode
            accessVlanId             = $vlan
            nativeVlanId             = $native
            allowedVlanIds           = $allowed
        }
    }
    $rec.networkAdapters = $na

    # Checkpoints
    $snaps = @()
    try {
        foreach ($s in @(Get-VMSnapshot -VM $vm -ErrorAction Stop)) {
            if ($null -eq $s) { continue }
            $snaps += [ordered]@{
                name               = S $s.Name
                id                 = S $s.Id
                creationTime       = D $s.CreationTime
                snapshotType       = S $s.SnapshotType
                parentSnapshotName = S $s.ParentSnapshotName
                path               = S $s.Path
            }
        }
    } catch { Add-Warning ("VM '" + $vm.Name + "' checkpoints: " + $_.Exception.Message) }
    $rec.snapshots = $snaps

    # Guest data from KVP
    $kv = $kvp[$id]
    if ($kv) {
        $rec.guest = [ordered]@{
            osName                     = $kv['OSName']
            osVersion                  = $kv['OSVersion']
            osBuildNumber              = $kv['OSBuildNumber']
            fullyQualifiedDomainName   = $kv['FullyQualifiedDomainName']
            integrationServicesVersion = $kv['IntegrationServicesVersion']
        }
    }
    return $rec
}

$vmList = New-Object 'System.Collections.Generic.List[object]'
$i = 0
foreach ($vm in $allVms) {
    $i++
    Write-Step ('VM {0}/{1}: {2}' -f $i, $allVms.Count, $vm.Name)
    try { $vmList.Add((Get-VmRecord $vm)) } catch { Add-Warning ("VM '" + $vm.Name + "': " + $_.Exception.Message) }
}

# ---------------------------------------------------------------- output

$result = [ordered]@{
    schemaVersion = $SchemaVersion
    generator     = 'Collect-HyperV.ps1'
    scriptVersion = $ScriptVersion
    collectedAt   = (Get-Date).ToString('o')
    computerName  = $env:COMPUTERNAME
    psVersion     = [string]$PSVersionTable.PSVersion
    host          = $hostData
    cluster       = $clusterData
    vms           = $vmList.ToArray()
    warnings      = $Warnings.ToArray()
}

Write-Step ('Done: {0} VM(s), {1} warning(s)' -f $vmList.Count, $Warnings.Count)
$json = ConvertTo-Json -InputObject $result -Depth 12 -Compress
if ($OutFile) {
    [System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Step ('Wrote ' + $OutFile)
} else {
    $json
}
