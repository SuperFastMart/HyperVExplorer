using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using HypervisorExplorer.Core.Model;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.App.ViewModels;

public enum TreeNodeKind
{
    All,
    Platform,
    Datacenter,
    Cluster,
    Host,
}

/// <summary>A node in the inventory tree. Selecting it scopes every table to the rows beneath it.</summary>
public sealed partial class TreeNodeViewModel : ObservableObject
{
    public required TreeNodeKind Kind { get; init; }
    public required string Key { get; init; }
    public required string Title { get; init; }
    public string? Subtitle { get; init; }
    public string Glyph => Kind switch
    {
        TreeNodeKind.All => "◉",
        TreeNodeKind.Platform => "▣",
        TreeNodeKind.Datacenter => "⌂",
        TreeNodeKind.Cluster => "⬡",
        _ => "▪",
    };

    public ObservableCollection<TreeNodeViewModel> Children { get; } = [];

    [ObservableProperty] private bool _isExpanded = true;

    public HashSet<string> Sources { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public HashSet<string> Clusters { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    /// <summary>(source, host) pairs under this node.</summary>
    public HashSet<(string, string)> Hosts { get; init; } = [];

    public bool Matches(GridRow row)
    {
        var (src, host, cluster) = row.Scope;
        if (string.IsNullOrEmpty(host)) host = null;
        if (string.IsNullOrEmpty(cluster)) cluster = null;

        if (Kind == TreeNodeKind.All) return true;
        // Rows not tied to any source (export metadata) apply everywhere.
        if (string.IsNullOrEmpty(src)) return true;
        if (!Sources.Contains(src)) return false;
        if (Kind == TreeNodeKind.Platform) return true;
        if (host is not null) return Hosts.Contains((src, host));
        if (cluster is not null) return Clusters.Contains(cluster);
        // Source-level rows (vSource, vHealth for the source, licences) belong to every node under that source.
        return true;
    }

    public static TreeNodeViewModel Build(Inventory inv)
    {
        var root = new TreeNodeViewModel
        {
            Kind = TreeNodeKind.All, Key = "all", Title = "All sources",
            Subtitle = $"{inv.Hosts.Count} hosts · {inv.VirtualMachines.Count} VMs",
        };

        int VmCount(HostSystem h) => inv.VirtualMachines.Count(v => v.SourceAddress == h.SourceAddress && v.Host == h.Name);

        foreach (var pg in inv.Hosts.GroupBy(h => h.Platform).OrderBy(g => g.Key))
        {
            var platformSources = pg.Select(h => h.SourceAddress).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var pNode = new TreeNodeViewModel
            {
                Kind = TreeNodeKind.Platform, Key = $"p:{pg.Key}", Title = RvToolsTables.PlatformName(pg.Key),
                Subtitle = $"{pg.Count()} hosts", Sources = platformSources,
            };
            root.Children.Add(pNode);

            foreach (var dg in pg.GroupBy(h => h.Datacenter ?? inv.FindSource(h.SourceAddress)?.Group ?? h.SourceAddress)
                         .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
            {
                var dNode = new TreeNodeViewModel
                {
                    Kind = TreeNodeKind.Datacenter, Key = $"d:{pg.Key}:{dg.Key}", Title = dg.Key,
                    Subtitle = $"{dg.Count()} hosts",
                    Sources = dg.Select(h => h.SourceAddress).ToHashSet(StringComparer.OrdinalIgnoreCase),
                    Clusters = dg.Where(h => h.Cluster is not null).Select(h => h.Cluster!).ToHashSet(StringComparer.OrdinalIgnoreCase),
                    Hosts = dg.Select(h => (h.SourceAddress, h.Name)).ToHashSet(),
                };
                pNode.Children.Add(dNode);

                foreach (var cg in dg.GroupBy(h => h.Cluster).OrderBy(g => g.Key is null).ThenBy(g => g.Key))
                {
                    var parent = dNode;
                    if (cg.Key is not null)
                    {
                        parent = new TreeNodeViewModel
                        {
                            Kind = TreeNodeKind.Cluster, Key = $"c:{cg.First().SourceAddress}:{cg.Key}", Title = cg.Key,
                            Subtitle = $"{cg.Count()} hosts · {cg.Sum(VmCount)} VMs",
                            Sources = cg.Select(h => h.SourceAddress).ToHashSet(StringComparer.OrdinalIgnoreCase),
                            Clusters = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { cg.Key },
                            Hosts = cg.Select(h => (h.SourceAddress, h.Name)).ToHashSet(),
                        };
                        dNode.Children.Add(parent);
                    }

                    foreach (var h in cg.OrderBy(h => h.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        parent.Children.Add(new TreeNodeViewModel
                        {
                            Kind = TreeNodeKind.Host, Key = $"h:{h.SourceAddress}:{h.Name}", Title = h.Name,
                            Subtitle = $"{VmCount(h)} VMs", Sources = { h.SourceAddress }, Hosts = { (h.SourceAddress, h.Name) },
                            Clusters = h.Cluster is null
                                ? new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                                : new HashSet<string>(StringComparer.OrdinalIgnoreCase) { h.Cluster },
                        });
                    }
                }
            }
        }
        return root;
    }

    public IEnumerable<TreeNodeViewModel> SelfAndDescendants()
    {
        yield return this;
        foreach (var c in Children)
        foreach (var d in c.SelfAndDescendants())
            yield return d;
    }
}
