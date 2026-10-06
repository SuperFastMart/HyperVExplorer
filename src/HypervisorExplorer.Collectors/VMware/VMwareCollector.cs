using System.Globalization;
using System.Xml.Linq;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Collectors.VMware;

/// <summary>Raw property data retrieved from one vSphere endpoint, before mapping.</summary>
public sealed class VSphereData
{
    public required ServiceContent Content { get; init; }
    /// <summary>Every ManagedEntity (name, parent): folders, datacenters, clusters, pools, networks, VMs, hosts…</summary>
    public List<VimObject> Entities { get; init; } = [];
    public List<VimObject> VirtualMachines { get; init; } = [];
    public List<VimObject> Hosts { get; init; } = [];
    public List<VimObject> Datastores { get; init; } = [];
    public List<VimObject> ComputeResources { get; init; } = [];
    public List<VimObject> ResourcePools { get; init; } = [];
    public List<VimObject> DistributedSwitches { get; init; } = [];
    public List<VimObject> DistributedPortgroups { get; init; } = [];
    /// <summary>LicenseManager.licenses (LicenseManagerLicenseInfo elements).</summary>
    public List<XElement> Licenses { get; init; } = [];
    /// <summary>LicenseAssignmentManager.QueryAssignedLicenses results (vCenter only).</summary>
    public List<XElement> LicenseAssignments { get; init; } = [];
}

/// <summary>
/// Collects VMware ESXi / vCenter inventory through the vSphere Web Services (SOAP) API at https://host/sdk,
/// like RVTools: Login → ContainerView per object type → paged RetrievePropertiesEx → Logout.
/// </summary>
public sealed class VMwareCollector : IInventoryCollector
{
    private readonly Func<HttpMessageHandler>? _handlerFactory;

    /// <param name="handlerFactory">Optional transport override (tests); default is a SocketsHttpHandler.</param>
    public VMwareCollector(Func<HttpMessageHandler>? handlerFactory = null) => _handlerFactory = handlerFactory;

    public Platform Platform => Platform.VMware;

    /// <summary>Objects per RetrievePropertiesEx page.</summary>
    public int PageSize { get; init; } = 500;

    internal static readonly string[] EntityPaths = ["name", "parent"];

    internal static readonly string[] VmPaths =
    [
        "name", "parent", "parentVApp", "resourcePool", "datastore", "network", "overallStatus", "configStatus",
        "guestHeartbeatStatus",
        "config.template", "config.uuid", "config.instanceUuid", "config.guestFullName", "config.guestId", "config.version",
        "config.changeVersion", "config.createDate", "config.annotation", "config.firmware", "config.bootOptions",
        "config.hardware", "config.cpuAllocation", "config.memoryAllocation", "config.cpuHotAddEnabled",
        "config.cpuHotRemoveEnabled", "config.memoryHotAddEnabled", "config.memoryReservationLockedToMax",
        "config.latencySensitivity", "config.changeTrackingEnabled", "config.files", "config.ftInfo", "config.managedBy",
        "config.tools", "config.scheduledHardwareUpgradeInfo", "config.extraConfig[\"disk.EnableUUID\"]",
        "runtime.powerState", "runtime.connectionState", "runtime.host", "runtime.bootTime", "runtime.consolidationNeeded",
        "runtime.faultToleranceState", "runtime.suspendTime", "runtime.suspendInterval", "runtime.minRequiredEVCModeKey",
        "runtime.dasVmProtection", "runtime.maxCpuUsage", "runtime.maxMemoryUsage", "runtime.memoryOverhead",
        "guest", "summary.storage", "summary.quickStats", "snapshot",
        "layoutEx.file", "layoutEx.disk", "layoutEx.snapshot",
    ];

    internal static readonly string[] HostPaths =
    [
        "name", "parent", "overallStatus", "configStatus",
        "runtime.connectionState", "runtime.inMaintenanceMode", "runtime.inQuarantineMode", "runtime.bootTime",
        "runtime.powerState",
        "summary.hardware", "summary.quickStats", "summary.config.product", "summary.currentEVCModeKey",
        "summary.maxEVCModeKey",
        "hardware.biosInfo", "hardware.systemInfo", "hardware.cpuPowerManagementInfo",
        "capability.vmotionSupported", "capability.storageVMotionSupported",
        "config.hyperThread", "config.network", "config.storageDevice.hostBusAdapter",
        "config.storageDevice.multipathInfo", "config.storageDevice.scsiLun", "config.fileSystemVolume.mountInfo",
        "config.dateTimeInfo", "config.powerSystemInfo", "config.service",
    ];

    internal static readonly string[] DatastorePaths =
        ["name", "parent", "overallStatus", "configStatus", "summary", "info", "host", "vm", "iormConfiguration"];

    internal static readonly string[] ComputeResourcePaths =
        ["name", "parent", "overallStatus", "configStatus", "summary", "configurationEx", "host", "resourcePool"];

    internal static readonly string[] ResourcePoolPaths =
        ["name", "parent", "owner", "vm", "resourcePool", "summary", "overallStatus"];

    internal static readonly string[] DvsPaths = ["name", "parent", "uuid", "overallStatus", "summary", "config"];

    internal static readonly string[] PortgroupPaths = ["name", "key", "config"];

    public async Task<InventorySnapshot> CollectAsync(ConnectionRequest request, IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.Username) || request.Secret is null)
            throw new CollectionException(CollectionFailure.AuthenticationFailed, "A user name and password are required for vSphere.",
                "Use a vCenter SSO account like administrator@vsphere.local or an ESXi local account (root); a read-only role is sufficient.");

        using var client = VSphereClient.Create(request, _handlerFactory?.Invoke());
        var warnings = new List<string>();
        var loggedIn = false;
        try
        {
            progress?.Report($"Connecting to {client.SdkUri}…");
            var sc = await client.RetrieveServiceContentAsync(cancellationToken);
            progress?.Report($"Connected to {sc.FullName}. Signing in as {request.Username}…");
            await client.LoginAsync(request.Username!, request.Secret, cancellationToken);
            loggedIn = true;

            var data = await RetrieveAllAsync(client, sc, progress, warnings, cancellationToken);
            progress?.Report("Building inventory…");
            var snapshot = new VMwareInventoryMapper(request, data).Map();
            snapshot.Warnings.AddRange(warnings);
            progress?.Report($"Collected {snapshot.VirtualMachines.Count:N0} VMs, {snapshot.Hosts.Count:N0} hosts, "
                + $"{snapshot.Datastores.Count:N0} datastores from {sc.Name}.");
            return snapshot;
        }
        catch (VimFaultException f)
        {
            throw VSphereClient.ToCollectionException(f, request.Username);
        }
        finally
        {
            if (loggedIn)
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
                try { await client.LogoutAsync(cts.Token); }
                catch (Exception) { /* best effort */ }
            }
        }
    }

    /// <summary>Retrieves every object type the mapper needs.</summary>
    public async Task<VSphereData> RetrieveAllAsync(VSphereClient client, ServiceContent sc, IProgress<string>? progress,
        List<string> warnings, CancellationToken ct)
    {
        var data = new VSphereData { Content = sc };
        data.Entities.AddRange(await RetrieveTypeAsync(client, sc, "ManagedEntity", "inventory objects", EntityPaths, progress, warnings, ct));
        data.Hosts.AddRange(await RetrieveTypeAsync(client, sc, "HostSystem", "hosts", HostPaths, progress, warnings, ct));
        data.ComputeResources.AddRange(await RetrieveTypeAsync(client, sc, "ComputeResource", "clusters / compute resources", ComputeResourcePaths, progress, warnings, ct));
        data.ResourcePools.AddRange(await RetrieveTypeAsync(client, sc, "ResourcePool", "resource pools", ResourcePoolPaths, progress, warnings, ct));
        data.Datastores.AddRange(await RetrieveTypeAsync(client, sc, "Datastore", "datastores", DatastorePaths, progress, warnings, ct));
        if (sc.IsVCenter)
        {
            data.DistributedSwitches.AddRange(await OptionalAsync(() =>
                RetrieveTypeAsync(client, sc, "DistributedVirtualSwitch", "distributed switches", DvsPaths, progress, warnings, ct), "distributed switches", warnings));
            data.DistributedPortgroups.AddRange(await OptionalAsync(() =>
                RetrieveTypeAsync(client, sc, "DistributedVirtualPortgroup", "distributed port groups", PortgroupPaths, progress, warnings, ct), "distributed port groups", warnings));
        }
        data.VirtualMachines.AddRange(await RetrieveTypeAsync(client, sc, "VirtualMachine", "VMs", VmPaths, progress, warnings, ct));

        if (sc.LicenseManager is { } lm)
        {
            try
            {
                progress?.Report("Retrieving licences…");
                var lic = await client.RetrieveObjectAsync(lm, ["licenses", "licenseAssignmentManager"], warnings, ct);
                if (lic is not null)
                {
                    data.Licenses.AddRange(lic.Items("licenses"));
                    if (lic.Ref1("licenseAssignmentManager") is { } lam)
                    {
                        var resp = await client.InvokeAsync("QueryAssignedLicenses", lam.ToXml("_this"), ct);
                        data.LicenseAssignments.AddRange(resp.Els("returnval"));
                    }
                }
            }
            catch (VimFaultException f)
            {
                warnings.Add($"Licences could not be read ({f.FaultType}: {f.Message}).");
            }
        }
        return data;
    }

    private static async Task<List<VimObject>> OptionalAsync(Func<Task<List<VimObject>>> call, string what, List<string> warnings)
    {
        try { return await call(); }
        catch (VimFaultException f)
        {
            warnings.Add($"Could not retrieve {what} ({f.FaultType}: {f.Message}).");
            return [];
        }
    }

    private async Task<List<VimObject>> RetrieveTypeAsync(VSphereClient client, ServiceContent sc, string type, string label,
        IReadOnlyCollection<string> paths, IProgress<string>? progress, List<string> warnings, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        progress?.Report($"Retrieving {label}…");
        var view = await client.CreateContainerViewAsync(sc.RootFolder, [type], true, ct);
        try
        {
            var list = await client.RetrieveFromViewAsync(view, type, paths, PageSize,
                n => { if (n >= PageSize) progress?.Report($"Retrieving {label}… {n.ToString("N0", CultureInfo.InvariantCulture)} so far"); },
                warnings, ct);
            progress?.Report($"Retrieved {list.Count.ToString("N0", CultureInfo.InvariantCulture)} {label}.");
            return list;
        }
        finally
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
            try { await client.DestroyViewAsync(view, cts.Token); }
            catch (Exception) { /* views die with the session anyway */ }
        }
    }
}
