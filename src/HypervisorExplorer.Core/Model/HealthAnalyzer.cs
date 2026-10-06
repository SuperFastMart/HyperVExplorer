namespace HypervisorExplorer.Core.Model;

/// <summary>Rule-based health checks, in the spirit of RVTools' vHealth tab.</summary>
public static class HealthAnalyzer
{
    public static int SnapshotAgeWarningDays { get; set; } = 7;
    public static double DatastoreFreeWarningPercent { get; set; } = 15;
    public static double DatastoreFreeErrorPercent { get; set; } = 5;

    public static List<HealthItem> Analyze(Inventory inv)
    {
        var items = new List<HealthItem>();
        var now = DateTimeOffset.Now;

        foreach (var vm in inv.VirtualMachines)
        {
            foreach (var snap in vm.Snapshots)
            {
                if (snap.Created is { } created && (now - created).TotalDays > SnapshotAgeWarningDays)
                {
                    items.Add(new HealthItem
                    {
                        Name = vm.Name,
                        Message = $"Snapshot '{snap.Name}' is {(int)(now - created).TotalDays} days old",
                        Severity = HealthSeverity.Warning,
                        SourceAddress = vm.SourceAddress,
                    });
                }
            }

            if (vm.PowerState == PowerState.PoweredOn && vm.Kind == GuestKind.VirtualMachine && vm.ToolsRunning == false)
            {
                var tools = vm.Platform switch
                {
                    Platform.HyperV => "Integration services",
                    Platform.Proxmox => "QEMU guest agent",
                    _ => "VMware Tools",
                };
                items.Add(new HealthItem
                {
                    Name = vm.Name,
                    Message = $"{tools} not running",
                    Severity = HealthSeverity.Info,
                    SourceAddress = vm.SourceAddress,
                });
            }

            foreach (var cd in vm.CdDrives.Where(c => c.Connected == true && !string.IsNullOrWhiteSpace(c.Media)))
            {
                items.Add(new HealthItem
                {
                    Name = vm.Name,
                    Message = $"Media mounted on {cd.DeviceNode}: {cd.Media}",
                    Severity = HealthSeverity.Info,
                    SourceAddress = vm.SourceAddress,
                });
            }
        }

        foreach (var ds in inv.Datastores)
        {
            if (ds.CapacityBytes is > 0 && ds.FreeBytes is { } free)
            {
                var pct = free * 100.0 / ds.CapacityBytes.Value;
                if (pct < DatastoreFreeWarningPercent)
                {
                    items.Add(new HealthItem
                    {
                        Name = ds.Name,
                        Message = $"Datastore has {pct:0.#}% free space",
                        Severity = pct < DatastoreFreeErrorPercent ? HealthSeverity.Error : HealthSeverity.Warning,
                        SourceAddress = ds.SourceAddress,
                    });
                }
            }
            if (!ds.Accessible)
            {
                items.Add(new HealthItem
                {
                    Name = ds.Name,
                    Message = "Datastore is not accessible",
                    Severity = HealthSeverity.Error,
                    SourceAddress = ds.SourceAddress,
                });
            }
        }

        foreach (var host in inv.Hosts.Where(h => h.InMaintenance))
        {
            items.Add(new HealthItem
            {
                Name = host.Name,
                Message = "Host is in maintenance mode",
                Severity = HealthSeverity.Warning,
                SourceAddress = host.SourceAddress,
            });
        }

        return items;
    }
}
