using System.Xml.Linq;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Collectors.VMware;

public sealed partial class VMwareInventoryMapper
{
    private sealed record Controller(string Label, string Type, int? Bus, string? SharedBus, bool IsScsi, XElement Element);

    private sealed record LayoutFile(string Name, string Type, long Size);

    private static readonly HashSet<string> EthernetTypes = new(StringComparer.Ordinal)
    {
        "VirtualEthernetCard", "VirtualVmxnet", "VirtualVmxnet2", "VirtualVmxnet3", "VirtualVmxnet3Vrdma",
        "VirtualE1000", "VirtualE1000e", "VirtualPCNet32", "VirtualSriovEthernetCard",
    };

    private void MapVirtualMachines()
    {
        foreach (var o in _d.VirtualMachines)
        {
            try
            {
                _snap.VirtualMachines.Add(MapVm(o));
            }
            catch (Exception ex)
            {
                _snap.Warnings.Add($"VM '{o.Str("name") ?? o.Ref.Value}' could not be mapped: {ex.Message}");
            }
        }
    }

    internal static PowerState MapPower(string? s) => s switch
    {
        "poweredOn" => PowerState.PoweredOn,
        "poweredOff" => PowerState.PoweredOff,
        "suspended" => PowerState.Suspended,
        _ => PowerState.Unknown,
    };

    /// <summary>"vmx-19" → 19.</summary>
    internal static int? HwVersionNumber(string? v) =>
        v is not null && v.StartsWith("vmx-", StringComparison.Ordinal) && int.TryParse(v[4..], out var n) ? n : null;

    private VirtualMachine MapVm(VimObject o)
    {
        var hostRef = o.Ref1("runtime.host");
        var parent = o.Ref1("parent");
        var vapp = o.Ref1("parentVApp");
        var folderRef = parent ?? vapp;
        var hw = o.Get("config.hardware");
        var numCpu = hw.Int("numCPU") ?? 0;
        var cores = hw.Int("numCoresPerSocket");
        var cpuAlloc = o.Get("config.cpuAllocation");
        var memAlloc = o.Get("config.memoryAllocation");
        var qs = o.Get("summary.quickStats");
        var guest = o.Get("guest");
        var toolsRunning = guest.Str("toolsRunningStatus");
        var hwVersion = o.Str("config.version");

        var vm = new VirtualMachine
        {
            Name = o.Str("name") ?? NameOf(o.Ref),
            VmId = o.Ref.Value,
            Uuid = o.Str("config.instanceUuid"),
            BiosUuid = o.Str("config.uuid"),
            Kind = GuestKind.VirtualMachine,
            IsTemplate = o.Bool("config.template") ?? false,
            Platform = Platform.VMware,
            Host = hostRef is null ? "" : NameOf(hostRef),
            Cluster = ClusterOfHost(hostRef),
            Datacenter = DatacenterOf(folderRef) ?? DatacenterOf(hostRef),
            SourceAddress = _addr,
            SourceProduct = _sc.Name,
            SourceApiVersion = _sc.ApiVersion,
            ResourcePool = InventoryPath(folderRef),
            PowerState = MapPower(o.Str("runtime.powerState")),
            RawState = o.Str("runtime.powerState"),
            Status = o.Str("overallStatus"),
            Heartbeat = o.Str("guestHeartbeatStatus"),
            DnsName = guest.Str("hostName"),
            CreationDate = o.Date("config.createDate"),
            PowerOnTime = o.Date("runtime.bootTime"),
            Uptime = qs.Long("uptimeSeconds") is { } up && o.Str("runtime.powerState") == "poweredOn" ? TimeSpan.FromSeconds(up) : null,
            CpuCount = numCpu,
            CoresPerSocket = cores,
            Sockets = cores is > 0 ? numCpu / cores.Value : null,
            CpuLimit = (int?)cpuAlloc.Long("limit"),
            CpuReservation = (int?)cpuAlloc.Long("reservation"),
            CpuShares = cpuAlloc.Int("shares.shares"),
            CpuHotAdd = o.Bool("config.cpuHotAddEnabled"),
            MemoryMiB = hw.Long("memoryMB") ?? 0,
            MemoryAssignedMiB = qs.Long("hostMemoryUsage"),
            MemoryDemandMiB = qs.Long("guestMemoryUsage"),
            MemoryMinMiB = memAlloc.Long("reservation"),
            MemoryMaxMiB = o.Long("runtime.maxMemoryUsage"),
            // The vMemory sheet's "Hot Add" column reads DynamicMemory.
            DynamicMemory = o.Bool("config.memoryHotAddEnabled"),
            MemoryBalloonedMiB = qs.Long("balloonedMemory"),
            Firmware = o.Str("config.firmware"),
            SecureBoot = o.Bool("config.bootOptions.efiSecureBootEnabled"),
            HardwareVersion = hwVersion,
            OsConfigured = o.Str("config.guestFullName"),
            OsGuest = guest.Str("guestFullName"),
            Annotation = o.Str("config.annotation"),
            ConfigPath = o.Str("config.files.vmPathName"),
            SnapshotDirectory = o.Str("config.files.snapshotDirectory"),
            HaProtected = o.Bool("runtime.dasVmProtection.dasProtected"),
            ToolsStatus = guest.Str("toolsStatus") ?? guest.Str("toolsVersionStatus2"),
            ToolsVersion = guest.Str("toolsVersion"),
            ToolsRunning = toolsRunning is null ? null : toolsRunning == "guestToolsRunning",
        };

        var x = vm.Extra;
        // vInfo
        x["Config status"] = o.Str("configStatus");
        x["Connection state"] = o.Str("runtime.connectionState");
        x["Guest state"] = guest.Str("guestState");
        x["Consolidation Needed"] = o.Bool("runtime.consolidationNeeded");
        x["Suspend time"] = o.Date("runtime.suspendTime");
        x["Suspend Interval"] = o.Long("runtime.suspendInterval");
        x["Change Version"] = o.Str("config.changeVersion");
        x["min Required EVC Mode Key"] = o.Str("runtime.minRequiredEVCModeKey");
        x["Latency Sensitivity"] = o.Str("config.latencySensitivity.level");
        x["EnableUUID"] = ExtraConfigValue(o, "disk.EnableUUID") is { } eu ? string.Equals(eu, "TRUE", StringComparison.OrdinalIgnoreCase) : null;
        x["CBT"] = o.Bool("config.changeTrackingEnabled");
        x["Primary IP Address"] = guest.Str("ipAddress");
        x["Resource pool"] = InventoryPath(o.Ref1("resourcePool"));
        x["Folder ID"] = parent?.Value;
        x["Folder"] = vm.ResourcePool;
        x["vApp"] = vapp is null ? null : NameOf(vapp);
        x["DAS protection"] = o.Bool("runtime.dasVmProtection.dasProtected");
        x["FT State"] = o.Str("runtime.faultToleranceState");
        x["FT Role"] = o.Int("config.ftInfo.role");
        var committed = o.Long("summary.storage.committed");
        var uncommitted = o.Long("summary.storage.uncommitted");
        if (committed is not null)
        {
            x["Provisioned MiB"] = Mib(committed + (uncommitted ?? 0));
            x["In Use MiB"] = Mib(committed);
            x["Unshared MiB"] = Mib(o.Long("summary.storage.unshared"));
        }
        x["Boot delay"] = o.Long("config.bootOptions.bootDelay");
        x["Boot retry delay"] = o.Long("config.bootOptions.bootRetryDelay");
        x["Boot retry enabled"] = o.Bool("config.bootOptions.bootRetryEnabled");
        x["Boot BIOS setup"] = o.Bool("config.bootOptions.enterBIOSSetup");
        x["EFI Secure boot"] = o.Bool("config.bootOptions.efiSecureBootEnabled");
        x["HW version"] = HwVersionNumber(hwVersion) is { } hv ? hv : hwVersion;
        x["HW upgrade status"] = o.Str("config.scheduledHardwareUpgradeInfo.scheduledHardwareUpgradeStatus");
        x["HW upgrade policy"] = o.Str("config.scheduledHardwareUpgradeInfo.upgradePolicy");
        x["HW target"] = o.Str("config.scheduledHardwareUpgradeInfo.versionKey");
        x["Log directory"] = o.Str("config.files.logDirectory");
        x["Suspend directory"] = o.Str("config.files.suspendDirectory");
        x["Customization Info"] = guest.Str("customizationInfo.customizationStatus");
        x["Guest Detailed Data"] = guest.Str("guestDetailedData");
        x["SRM Placeholder"] = o.Str("config.managedBy.extensionKey") == "com.vmware.vcDr"
            && o.Str("config.managedBy.type") == "placeholderVm";
        x["VI SDK Server type"] = _sc.Name;
        x["VI SDK API Version"] = _sc.ApiVersion;
        ApplyClusterHa(vm, o, hostRef);

        // vCPU (headers shared with vMemory are sheet-qualified)
        x["vCPU:Level"] = cpuAlloc.Str("shares.level");
        x["vCPU:Max"] = o.Long("runtime.maxCpuUsage");
        x["Overall"] = qs.Long("overallCpuUsage");
        x["vCPU:Entitlement"] = qs.Long("staticCpuEntitlement");
        x["vCPU:DRS Entitlement"] = qs.Long("distributedCpuEntitlement");
        x["Hot Remove"] = o.Bool("config.cpuHotRemoveEnabled");

        // vMemory
        x["Memory Reservation Locked To Max"] = o.Bool("config.memoryReservationLockedToMax");
        x["Overhead"] = Mib(o.Long("runtime.memoryOverhead"));
        x["Consumed Overhead"] = qs.Long("consumedOverheadMemory");
        x["Private"] = qs.Long("privateMemory");
        x["Shared"] = qs.Long("sharedMemory");
        x["Swapped"] = qs.Long("swappedMemory");
        x["vMemory:Entitlement"] = qs.Long("staticMemoryEntitlement");
        x["vMemory:DRS Entitlement"] = qs.Long("distributedMemoryEntitlement");
        x["vMemory:Level"] = memAlloc.Str("shares.level");
        x["vMemory:Shares"] = memAlloc.Int("shares.shares");
        x["vMemory:Limit"] = memAlloc.Long("limit");

        // vTools
        var tvs = guest.Str("toolsVersionStatus2");
        x["VM Version"] = HwVersionNumber(hwVersion) is { } hv2 ? hv2 : hwVersion;
        x["Upgradeable"] = tvs is null ? null : tvs is "guestToolsNeedUpgrade" or "guestToolsSupportedOld" or "guestToolsTooOld";
        x["Upgrade Policy"] = o.Str("config.tools.toolsUpgradePolicy");
        x["Sync time"] = o.Bool("config.tools.syncTimeWithHost");
        x["App status"] = guest.Str("appHeartbeatStatus");
        x["Heartbeat status"] = o.Str("guestHeartbeatStatus");
        x["Kernel Crash state"] = guest.Str("guestKernelCrashed");
        x["Operation Ready"] = guest.Bool("guestOperationsReady");
        x["State change support"] = guest.Bool("guestStateChangeSupported");
        x["Interactive Guest"] = guest.Bool("interactiveGuestOperationsReady");

        var files = LayoutFiles(o);
        MapDevices(vm, o, hw, hostRef, files);
        MapSnapshots(vm, o, files);
        foreach (var gd in guest.Els("disk"))
        {
            var p = new VmPartition
            {
                Name = gd.Str("diskPath") ?? "",
                FileSystem = gd.Str("filesystemType"),
                CapacityBytes = gd.Long("capacity"),
                FreeBytes = gd.Long("freeSpace"),
            };
            p.Extra["Disk Key"] = Join(gd.Els("mappings").Select(m => m.Str("key")));
            vm.Partitions.Add(p);
        }
        return vm;
    }

    private static string? ExtraConfigValue(VimObject o, string key)
    {
        foreach (var (path, val) in o.Props)
        {
            if (!path.StartsWith("config.extraConfig", StringComparison.Ordinal)) continue;
            var options = val.XsiType() == "OptionValue" ? [val] : val.Elements();
            foreach (var ov in options)
            {
                if (string.Equals(ov.Str("key"), key, StringComparison.OrdinalIgnoreCase)) return ov.Str("value");
            }
        }
        return null;
    }

    private void ApplyClusterHa(VirtualMachine vm, VimObject o, MoRef? hostRef)
    {
        if (ComputeResourceOfHost(hostRef) is not { Type: "ClusterComputeResource" } c
            || !_clusterHa.TryGetValue(c.Value, out var ha))
            return;
        var x = vm.Extra;
        var rules = ha.Rules.Where(r => r.Vms.Contains(o.Ref.Value)).ToList();
        x["Cluster rule(s)"] = rules.Count == 0 ? null : string.Join(", ", rules.Select(r => r.Type));
        x["Cluster rule name(s)"] = rules.Count == 0 ? null : string.Join(", ", rules.Select(r => r.Name));
        if (!ha.HaEnabled)
        {
            x["HA Restart Priority"] = null;
            return;
        }
        var defaults = ha.DasConfig.El("defaultVmSettings");
        ha.VmOverrides.TryGetValue(o.Ref.Value, out var ovr);
        var settings = ovr.El("dasSettings");
        var restart = settings.Str("restartPriority") ?? ovr.Str("restartPriority");
        if (restart is null or "clusterRestartPriority") restart = defaults.Str("restartPriority");
        var isolation = settings.Str("isolationResponse");
        if (isolation is null or "clusterIsolationResponse") isolation = defaults.Str("isolationResponse");
        var monitoring = settings.Str("vmToolsMonitoringSettings.vmMonitoring");
        if (monitoring is null || settings.Bool("vmToolsMonitoringSettings.clusterSettings") == true)
            monitoring = defaults.Str("vmToolsMonitoringSettings.vmMonitoring") ?? ha.DasConfig.Str("vmMonitoring");
        x["HA Restart Priority"] = restart;
        x["HA Isolation Response"] = isolation;
        x["HA VM Monitoring"] = monitoring;
        vm.HaState = restart;
    }

    private static Dictionary<int, LayoutFile> LayoutFiles(VimObject o)
    {
        var files = new Dictionary<int, LayoutFile>();
        foreach (var f in o.Items("layoutEx.file"))
        {
            if (f.Int("key") is { } k)
                files[k] = new LayoutFile(f.Str("name") ?? "", f.Str("type") ?? "", f.Long("size") ?? 0);
        }
        return files;
    }

    private static long? ChainSize(IEnumerable<XElement> units, Dictionary<int, LayoutFile> files)
    {
        long total = 0;
        var any = false;
        foreach (var unit in units)
        {
            foreach (var fk in unit.Els("fileKey"))
            {
                if (int.TryParse(fk.Value, out var k) && files.TryGetValue(k, out var f))
                {
                    total += f.Size;
                    any = true;
                }
            }
        }
        return any ? total : null;
    }

    private static string AdapterName(string type) => type switch
    {
        "VirtualVmxnet3" => "vmxnet3",
        "VirtualVmxnet3Vrdma" => "vmxnet3vrdma",
        "VirtualVmxnet2" => "vmxnet2",
        "VirtualVmxnet" => "vmxnet",
        "VirtualE1000e" => "e1000e",
        "VirtualE1000" => "e1000",
        "VirtualPCNet32" => "pcnet32",
        "VirtualSriovEthernetCard" => "sriov",
        _ => type.StartsWith("Virtual", StringComparison.Ordinal) ? type[7..].ToLowerInvariant() : type,
    };

    private static (string Type, bool Scsi)? ControllerKind(string? t) => t switch
    {
        "ParaVirtualSCSIController" => ("VMware Paravirtual", true),
        "VirtualLsiLogicController" => ("LSI Logic Parallel", true),
        "VirtualLsiLogicSASController" => ("LSI Logic SAS", true),
        "VirtualBusLogicController" => ("BusLogic Parallel", true),
        "VirtualSCSIController" => ("SCSI", true),
        "VirtualAHCIController" or "VirtualSATAController" => ("SATA AHCI", false),
        "VirtualIDEController" => ("IDE", false),
        "VirtualNVMEController" => ("NVMe", false),
        "VirtualNVDIMMController" => ("NVDIMM", false),
        "VirtualUSBController" or "VirtualUSBXHCIController" => ("USB", false),
        _ => null,
    };

    private void MapDevices(VirtualMachine vm, VimObject o, XElement? hw, MoRef? hostRef, Dictionary<int, LayoutFile> files)
    {
        var devices = hw.Els("device").ToList();
        var controllers = new Dictionary<int, Controller>();
        foreach (var dev in devices)
        {
            if (ControllerKind(dev.XsiType()) is { } kind && dev.Int("key") is { } key)
                controllers[key] = new Controller(dev.Str("deviceInfo.label") ?? kind.Type, kind.Type, dev.Int("busNumber"),
                    dev.Str("sharedBus"), kind.Scsi, dev);
        }

        var diskLayouts = new Dictionary<int, XElement>();
        foreach (var dl in o.Items("layoutEx.disk"))
        {
            if (dl.Int("key") is { } k) diskLayouts[k] = dl;
        }

        var guestNics = o.Get("guest").Els("net").ToList();
        var portgroups = hostRef is null ? null : HostPortgroupSwitches(hostRef.Value);

        int diskIdx = 0, nicIdx = 0;
        foreach (var dev in devices)
        {
            var type = dev.XsiType() ?? "";
            var key = dev.Int("key");
            var label = dev.Str("deviceInfo.label") ?? type;
            var backing = dev.El("backing");
            var btype = backing.XsiType() ?? "";
            Controller? ctrl = dev.Int("controllerKey") is { } ck ? controllers.GetValueOrDefault(ck) : null;

            if (type == "VirtualDisk")
            {
                var fileName = backing.Str("fileName");
                var rdm = btype.Contains("RawDiskMapping", StringComparison.Ordinal);
                var sharing = backing.Str("sharing");
                var disk = new VmDisk
                {
                    Index = diskIdx++,
                    Label = label,
                    Controller = ctrl?.Label,
                    ControllerType = ctrl?.Type,
                    ControllerNumber = ctrl?.Bus,
                    Unit = dev.Int("unitNumber"),
                    Path = fileName,
                    Datastore = DatastoreFromPath(fileName) ?? (MoRef.From(backing.El("datastore")) is { } dsr ? NameOf(dsr) : null),
                    CapacityBytes = dev.Long("capacityInBytes") ?? dev.Long("capacityInKB") * 1024,
                    UsedBytes = key is { } dk && diskLayouts.TryGetValue(dk, out var layout) ? ChainSize(layout.Els("chain"), files) : null,
                    Format = backing.Str("diskMode"),
                    Thin = backing.Bool("thinProvisioned") ?? (rdm ? null : false),
                    Shared = sharing is null ? null : sharing != "sharingNone",
                    Passthrough = rdm,
                    ParentPath = backing.Str("parent.fileName"),
                    Cache = backing.Bool("writeThrough") == true ? "writeThrough" : null,
                };
                var x = disk.Extra;
                x["Disk Key"] = key;
                x["Disk UUID"] = backing.Str("uuid");
                x["Disk Mode"] = backing.Str("diskMode");
                x["Sharing mode"] = sharing;
                x["Thin"] = disk.Thin;
                x["Eagerly Scrub"] = backing.Bool("eagerlyScrub");
                x["Split"] = backing.Bool("split");
                x["Write Through"] = backing.Bool("writeThrough");
                x["Level"] = dev.Str("storageIOAllocation.shares.level") ?? dev.Str("shares.level");
                x["Shares"] = dev.Int("storageIOAllocation.shares.shares") ?? dev.Int("shares.shares");
                x["Reservation"] = dev.Long("storageIOAllocation.reservation");
                x["Limit"] = dev.Long("storageIOAllocation.limit");
                x["SCSI Unit #"] = ctrl is { IsScsi: true } ? dev.Int("unitNumber") : null;
                x["Shared Bus"] = ctrl?.SharedBus;
                x["Raw"] = rdm;
                x["Raw LUN ID"] = rdm ? backing.Str("lunUuid") : null;
                x["Raw Comp. Mode"] = rdm ? backing.Str("compatibilityMode") : null;
                x["Path"] = rdm ? backing.Str("deviceName") : null;
                vm.Disks.Add(disk);
            }
            else if (EthernetTypes.Contains(type) || (dev.El("macAddress") is not null && dev.El("addressType") is not null))
            {
                var nic = new VmNic
                {
                    Index = nicIdx++,
                    Label = label,
                    AdapterType = AdapterName(type),
                    Mac = dev.Str("macAddress"),
                    Connected = dev.Bool("connectable.connected"),
                    StartsConnected = dev.Bool("connectable.startConnected"),
                };
                switch (btype)
                {
                    case "VirtualEthernetCardDistributedVirtualPortBackingInfo":
                        var pgKey = backing.Str("port.portgroupKey");
                        var swUuid = backing.Str("port.switchUuid");
                        nic.Network = pgKey is not null && _pgByKey.TryGetValue(pgKey, out var pg) ? pg.Str("name") ?? pgKey : pgKey;
                        nic.Switch = swUuid is not null && _dvsByUuid.TryGetValue(swUuid, out var dvs) ? dvs.Str("name") : null;
                        if (nic.Switch is null && pgKey is not null && _pgByKey.TryGetValue(pgKey, out var pg2)
                            && pg2.Ref1("config.distributedVirtualSwitch") is { } swRef)
                            nic.Switch = NameOf(swRef);
                        break;
                    case "VirtualEthernetCardOpaqueNetworkBackingInfo":
                        nic.Network = backing.Str("opaqueNetworkId");
                        nic.Switch = backing.Str("opaqueNetworkType");
                        break;
                    default:
                        nic.Network = backing.Str("deviceName") ?? (MoRef.From(backing.El("network")) is { } nr ? NameOf(nr) : null);
                        if (nic.Network is not null && portgroups is not null && portgroups.TryGetValue(nic.Network, out var sw))
                            nic.Switch = sw;
                        break;
                }
                var gn = guestNics.FirstOrDefault(g => key is not null && g.Int("deviceConfigId") == key)
                    ?? guestNics.FirstOrDefault(g => nic.Mac is not null && string.Equals(g.Str("macAddress"), nic.Mac, StringComparison.OrdinalIgnoreCase));
                if (gn is not null)
                {
                    var ips = gn.El("ipConfig").Els("ipAddress").Select(a => a.Str("ipAddress")).OfType<string>().ToList();
                    if (ips.Count == 0) ips = gn.Strs("ipAddress");
                    nic.Ipv4.AddRange(ips.Where(i => !i.Contains(':')));
                    nic.Ipv6.AddRange(ips.Where(i => i.Contains(':')));
                }
                nic.Extra["Type"] = dev.Str("addressType");
                nic.Extra["Direct Path IO"] = dev.Bool("uptCompatibilityEnabled") ?? false;
                vm.Nics.Add(nic);
            }
            else if (type == "VirtualCdrom")
            {
                var cd = new VmCdDrive
                {
                    DeviceNode = label,
                    Media = backing.Str("fileName") ?? backing.Str("deviceName"),
                    Connected = dev.Bool("connectable.connected"),
                    DeviceType = dev.Str("deviceInfo.summary"),
                };
                if (string.IsNullOrWhiteSpace(cd.Media)) cd.Media = null;
                cd.Extra["Starts Connected"] = dev.Bool("connectable.startConnected");
                vm.CdDrives.Add(cd);
            }
            else if (type == "VirtualUSB")
            {
                var usb = new VmUsbDevice
                {
                    DeviceNode = label,
                    DeviceType = dev.Str("deviceInfo.summary"),
                    Connected = dev.Bool("connected") ?? dev.Bool("connectable.connected"),
                };
                usb.Extra["Family"] = Join(dev.Strs("family"));
                usb.Extra["Speed"] = Join(dev.Strs("speed"));
                usb.Extra["EHCI enabled"] = ctrl?.Element.Bool("ehciEnabled");
                usb.Extra["Auto connect"] = ctrl?.Element.Bool("autoConnectDevices");
                usb.Extra["Bus number"] = ctrl?.Bus;
                usb.Extra["Unit number"] = dev.Int("unitNumber");
                vm.UsbDevices.Add(usb);
            }
            else if (type == "VirtualMachineVideoCard")
            {
                vm.Extra["Num Monitors"] = dev.Int("numDisplays");
                vm.Extra["Video Ram KiB"] = dev.Long("videoRamSizeInKB");
            }
        }
    }

    private Dictionary<string, string> HostPortgroupSwitches(string hostId)
    {
        if (_hostPortgroupSwitch.TryGetValue(hostId, out var map)) return map;
        map = new Dictionary<string, string>(StringComparer.Ordinal);
        if (_hostObjs.TryGetValue(hostId, out var h))
        {
            foreach (var pg in h.Get("config.network").Els("portgroup"))
            {
                if (pg.Str("spec.name") is { } n && pg.Str("spec.vswitchName") is { } s) map[n] = s;
            }
        }
        _hostPortgroupSwitch[hostId] = map;
        return map;
    }

    private void MapSnapshots(VirtualMachine vm, VimObject o, Dictionary<int, LayoutFile> files)
    {
        var current = o.Ref1("snapshot.currentSnapshot")?.Value;
        var layouts = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var sl in o.Items("layoutEx.snapshot"))
        {
            if (MoRef.From(sl.El("key")) is { } k) layouts[k.Value] = sl;
        }
        var currentDisks = o.Items("layoutEx.disk").ToList();

        void Walk(XElement tree, string? parentName)
        {
            var name = tree.Str("name") ?? "";
            var id = MoRef.From(tree.El("snapshot"))?.Value;
            var state = tree.Str("state");
            var children = tree.Els("childSnapshotList").ToList();
            var snap = new VmSnapshot
            {
                Name = name,
                Description = tree.Str("description"),
                Created = tree.Date("createTime"),
                Parent = parentName,
                Type = state,
                IncludesMemory = state == "poweredOn",
            };
            snap.Extra["State"] = state;
            snap.Extra["Quiesced"] = tree.Bool("quiesced");
            if (id is not null && layouts.TryGetValue(id, out var sl))
            {
                var vmsn = sl.Int("dataKey") is { } dk && files.TryGetValue(dk, out var df) ? df : null;
                var mem = sl.Int("memoryKey") is { } mk && mk >= 0 && files.TryGetValue(mk, out var mf) ? mf : null;
                snap.Path = vmsn?.Name;
                snap.Extra["Size MiB (vmsn)"] = vmsn is null ? null : Mib(vmsn.Size + (mem?.Size ?? 0));
                // Growth since the snapshot: the delta disks that follow this snapshot's chain in the next state.
                var next = id == current ? currentDisks
                    : children.Select(c => MoRef.From(c.El("snapshot"))?.Value).OfType<string>()
                        .Select(cid => layouts.GetValueOrDefault(cid)).FirstOrDefault(l => l is not null).Els("disk").ToList();
                long delta = 0;
                foreach (var disk in sl.Els("disk"))
                {
                    var depth = disk.Els("chain").Count();
                    var nextDisk = next.FirstOrDefault(d => d.Int("key") == disk.Int("key"));
                    var unit = nextDisk.Els("chain").Skip(depth).FirstOrDefault();
                    delta += unit is null ? 0 : ChainSize([unit], files) ?? 0;
                }
                snap.SizeBytes = (vmsn?.Size ?? 0) + (mem?.Size ?? 0) + delta;
            }
            vm.Snapshots.Add(snap);
            foreach (var c in children) Walk(c, name);
        }

        foreach (var root in o.Items("snapshot.rootSnapshotList")) Walk(root, null);
    }
}
