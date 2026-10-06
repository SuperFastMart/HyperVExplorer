using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Collectors;

/// <summary>Creates the collector for a platform.</summary>
public static class CollectorRegistry
{
    public static IInventoryCollector Create(Platform platform) => platform switch
    {
        _ => throw new NotSupportedException($"No collector registered for {platform}."),
    };

    /// <summary>Default port shown in the connect dialog.</summary>
    public static int DefaultPort(Platform platform, bool useSsl = false) => platform switch
    {
        Platform.HyperV => useSsl ? 5986 : 5985,
        Platform.Proxmox => 8006,
        Platform.VMware => 443,
        _ => 0,
    };
}
