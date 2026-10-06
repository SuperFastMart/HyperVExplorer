using System.Collections.ObjectModel;
using System.Collections.Specialized;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using HypervisorExplorer.App.Services;
using HypervisorExplorer.Collectors.HyperV;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Config;
using HypervisorExplorer.Core.Export;
using HypervisorExplorer.Core.Model;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.App.ViewModels;

public sealed partial class MainViewModel : ViewModelBase
{
    private static readonly (string Category, string[] Tables)[] Layout =
    [
        ("Overview", ["hvOverview"]),
        ("Virtual machines", ["vInfo", "vCPU", "vMemory", "vDisk", "vPartition", "vNetwork", "vCD", "vUSB", "vSnapshot", "vTools"]),
        ("Hosts", ["vHost", "vHBA", "vNIC", "vSwitch", "vPort", "vSC_VMK", "dvSwitch", "dvPort"]),
        ("Clusters & storage", ["vCluster", "hvClusterNodes", "hvClusterNetworks", "hvHA", "vRP", "vDatastore", "vMultiPath"]),
        ("Other", ["vHealth", "vSource", "vLicense", "vFileInfo", "vMetaData"]),
    ];

    private readonly InventoryStore _store = new();
    private readonly ConfigStore _config;
    private readonly DispatcherTimer _searchDebounce;
    private readonly Dictionary<string, ConnectDialogResult> _pendingSaves = new(StringComparer.OrdinalIgnoreCase);
    private Inventory _inventory = Inventory.Empty;

    public MainViewModel() : this(new ConfigStore())
    {
    }

    public MainViewModel(ConfigStore config)
    {
        _config = config;
        HealthAnalyzer.SnapshotAgeWarningDays = config.Config.Settings.SnapshotAgeWarningDays;
        _hideEmptyColumns = config.Config.Settings.HideEmptyColumns;
        _sourcesExpanded = config.Config.Settings.SourcesExpanded;

        Connections = new ConnectionManager(_store, config.Config.Settings.MaxParallelConnections);
        Connections.Succeeded += OnConnectionSucceeded;
        Connections.Failed += item =>
        {
            if (_config.Config.FindHost(item.Request.Address, item.Platform) is { } host)
            {
                host.LastError = item.Message;
                _config.Save();
            }
        };
        Connections.Connections.CollectionChanged += (_, _) =>
        {
            OnPropertyChanged(nameof(HasConnections));
            UpdateSourcesSummary();
        };
        Connections.Activity.CollectionChanged += OnActivityChanged;

        var all = RvToolsTables.All.Concat(ExtendedTables.All).ToDictionary(t => t.Name);
        foreach (var (category, names) in Layout)
        foreach (var name in names)
            AllTables.Add(new TableViewModel(all[name], category));

        _selectedTable = AllTables[0];
        _store.Changed += (_, _) => Dispatcher.UIThread.Post(OnInventoryChanged);

        _searchDebounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
        _searchDebounce.Tick += (_, _) =>
        {
            _searchDebounce.Stop();
            ApplyFilter();
        };

        RebuildVisibleTables();
        OnInventoryChanged();
    }

    public IDialogService? Dialogs { get; set; }
    public ConnectionManager Connections { get; }
    public ObservableCollection<TableViewModel> AllTables { get; } = [];
    public ObservableCollection<TableViewModel> VisibleTables { get; } = [];
    public ObservableCollection<TreeNodeViewModel> TreeRoots { get; } = [];

    /// <summary>Raised when the grid must rebuild its columns (table switched, data or column visibility changed).</summary>
    public event Action? ColumnsChanged;

    [ObservableProperty] private TableViewModel? _selectedTable;
    [ObservableProperty] private TreeNodeViewModel? _selectedNode;
    [ObservableProperty] private string _searchText = "";
    [ObservableProperty] private VmDetailsViewModel? _details;
    [ObservableProperty] private bool _showDetails = true;
    [ObservableProperty] private bool _showActivity;
    [ObservableProperty] private bool _showEmptyTables;
    [ObservableProperty] private bool _hideEmptyColumns;
    [ObservableProperty] private string _statusText = "Ready — connect to a host, open a saved inventory, or load demo data.";
    [ObservableProperty] private string _summaryText = "";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private bool _sourcesExpanded;
    [ObservableProperty] private string _sourcesSummary = "";
    [ObservableProperty] private bool _anySourceFailed;

    partial void OnSourcesExpandedChanged(bool value)
    {
        _config.Config.Settings.SourcesExpanded = value;
        _config.Save();
    }

    /// <summary>One-line status for the (possibly collapsed) Sources panel header.</summary>
    private void UpdateSourcesSummary()
    {
        var all = Connections.Connections;
        var ok = all.Count(c => c.Status is ConnectionStatus.Connected or ConnectionStatus.Imported);
        var failed = all.Count(c => c.Status == ConnectionStatus.Failed);
        var busy = all.Count(c => c.IsBusy);
        var parts = new List<string> { $"{all.Count} source{(all.Count == 1 ? "" : "s")}" };
        if (ok > 0) parts.Add($"{ok} OK");
        if (busy > 0) parts.Add($"{busy} collecting");
        if (failed > 0) parts.Add($"{failed} failed");
        SourcesSummary = string.Join(" · ", parts);
        AnySourceFailed = failed > 0;
    }

    public bool HasConnections => Connections.Connections.Count > 0;
    public bool HasData => _inventory.Sources.Count > 0;
    public string WindowTitle => "Hypervisor Explorer " + RvToolsTables.AppVersion;

    /// <summary>Per-table user-hidden columns (by header).</summary>
    public Dictionary<string, HashSet<string>> HiddenColumns { get; } = [];

    partial void OnSelectedTableChanged(TableViewModel? value)
    {
        if (value is null) return;
        value.EnsureBuilt();
        ApplyFilter();
        ColumnsChanged?.Invoke();
    }

    partial void OnSelectedNodeChanged(TreeNodeViewModel? value) => ApplyFilter();

    partial void OnSearchTextChanged(string value)
    {
        _searchDebounce.Stop();
        _searchDebounce.Start();
    }

    partial void OnShowEmptyTablesChanged(bool value) => RebuildVisibleTables();

    partial void OnHideEmptyColumnsChanged(bool value)
    {
        _config.Config.Settings.HideEmptyColumns = value;
        _config.Save();
        ColumnsChanged?.Invoke();
    }

    private void OnActivityChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (Connections.Activity.FirstOrDefault() is { } latest)
            StatusText = $"{latest.Source}: {latest.Message}";
        IsBusy = Connections.AnyBusy;
        UpdateSourcesSummary();
    }

    private void OnInventoryChanged()
    {
        _inventory = _store.Current;
        foreach (var t in AllTables) t.SetInventory(_inventory);

        var selectedKey = SelectedNode?.Key;
        TreeRoots.Clear();
        if (_inventory.Sources.Count > 0)
        {
            var root = TreeNodeViewModel.Build(_inventory);
            TreeRoots.Add(root);
            SelectedNode = root.SelfAndDescendants().FirstOrDefault(n => n.Key == selectedKey) ?? root;
        }
        else
        {
            SelectedNode = null;
        }

        RebuildVisibleTables();
        if (SelectedTable is { } table)
        {
            table.EnsureBuilt();
            ApplyFilter();
            ColumnsChanged?.Invoke();
        }

        if (Details is { } d)
            Details = _inventory.VirtualMachines.FirstOrDefault(v => v.SourceAddress == d.Vm.SourceAddress && v.Name == d.Vm.Name) is { } vm
                ? new VmDetailsViewModel(vm) : null;

        var vms = _inventory.VirtualMachines;
        SummaryText = _inventory.Sources.Count == 0
            ? "No data"
            : $"{_inventory.Sources.Count} source(s) · {_inventory.Clusters.Count} cluster(s) · {_inventory.Hosts.Count} host(s) · " +
              $"{vms.Count:N0} VM(s), {vms.Count(v => v.PowerState == PowerState.PoweredOn):N0} on · " +
              $"{vms.Sum(v => v.CpuCount):N0} vCPU · {vms.Sum(v => v.MemoryMiB) / 1024.0:N0} GiB vRAM · " +
              $"{vms.Sum(v => v.ProvisionedBytes) / 1099511627776.0:N1} TiB provisioned";
        OnPropertyChanged(nameof(HasData));
        IsBusy = Connections.AnyBusy;
        ExportCommandsChanged();
    }

    private void ExportCommandsChanged()
    {
        ExportXlsxCommand.NotifyCanExecuteChanged();
        ExportCsvCommand.NotifyCanExecuteChanged();
        ExportCsvZipCommand.NotifyCanExecuteChanged();
        ExportHtmlCommand.NotifyCanExecuteChanged();
        SaveInventoryCommand.NotifyCanExecuteChanged();
    }

    private void RebuildVisibleTables()
    {
        var keep = SelectedTable;
        var wanted = AllTables.Where(t => ShowEmptyTables || t.TotalCount > 0 || t.Name is "hvOverview" or "vInfo").ToList();
        if (wanted.SequenceEqual(VisibleTables)) return;
        VisibleTables.Clear();
        foreach (var t in wanted) VisibleTables.Add(t);
        SelectedTable = keep is not null && wanted.Contains(keep) ? keep : wanted.FirstOrDefault();
    }

    private void ApplyFilter()
    {
        if (SelectedTable is not { } table) return;
        var node = SelectedNode;
        table.ApplyFilter(SearchText, node is null || node.Kind == TreeNodeKind.All ? null : node.Matches);
    }

    /// <summary>Columns to show for the current table, honouring empty-column hiding and user choices.</summary>
    public IReadOnlyList<(int Index, TableColumn Column)> VisibleColumns()
    {
        if (SelectedTable is not { } t) return [];
        t.EnsureBuilt();
        var hidden = HiddenColumns.GetValueOrDefault(t.Name);
        return t.Definition.Columns
            .Select((c, i) => (i, c))
            .Where(x => (!HideEmptyColumns || t.TotalCount == 0 || x.i >= t.ColumnHasData.Length || t.ColumnHasData[x.i] || x.i == 0)
                        && (hidden is null || !hidden.Contains(x.c.Header)))
            .ToList();
    }

    public void SetColumnHidden(string header, bool hidden)
    {
        if (SelectedTable is not { } t) return;
        if (!HiddenColumns.TryGetValue(t.Name, out var set)) HiddenColumns[t.Name] = set = [];
        if (hidden) set.Add(header);
        else set.Remove(header);
        ColumnsChanged?.Invoke();
    }

    public void OnRowSelected(GridRow? row)
    {
        if (row?.Vm is { } vm) Details = new VmDetailsViewModel(vm);
    }

    // ---------------- Connections ----------------

    [RelayCommand]
    private async Task Connect()
    {
        if (Dialogs is null) return;
        var vm = new ConnectDialogViewModel(_config.Config.Groups);
        if (await Dialogs.ShowConnectDialogAsync(vm) is not { } result) return;
        StartConnection(result.Request, result);
    }

    /// <summary>Starts a collection without awaiting it, so commands stay enabled and sources run in parallel.</summary>
    private void StartConnection(ConnectionRequest request, ConnectDialogResult? fromDialog)
    {
        if (request.CredentialKind != CredentialKind.CurrentUser && (request.Username is null || request.Secret is null)
            && fromDialog?.GroupId is { } gid && _config.Config.Groups.FirstOrDefault(g => g.Id == gid) is { } group)
        {
            // Inherit credentials from the chosen group.
            request = request with
            {
                CredentialKind = group.CredentialKind,
                Username = group.Username,
                Secret = group.ProtectedSecret is null ? null : _config.Protector.Unprotect(group.ProtectedSecret),
            };
        }
        if (fromDialog is not null) _pendingSaves[request.Address] = fromDialog;
        _ = Connections.ConnectAsync(request);
    }

    private void OnConnectionSucceeded(ConnectionRequest request)
    {
        _pendingSaves.Remove(request.Address, out var dialog);
        var existing = _config.Config.FindHost(request.Address, request.Platform);
        var host = _config.Remember(request, dialog?.SaveCredentials ?? existing?.ProtectedSecret is not null, dialog?.GroupId);
        if (dialog?.DisplayName is not null) host.DisplayName = dialog.DisplayName;
        _config.Save();
    }

    /// <summary>
    /// Connects saved hosts, prompting only for those without usable credentials. Hosts that are already connected
    /// or collecting are skipped unless <paramref name="includeConnected"/> (used by Refresh). Returns how many started.
    /// </summary>
    public async Task<int> ConnectSaved(IReadOnlyList<SavedHost> hosts, bool includeConnected = false)
    {
        var started = 0;
        foreach (var host in hosts)
        {
            var request = _config.BuildRequest(host);
            if (!includeConnected && Connections.Find(request.DisplayName) is { } existing
                && (existing.IsBusy || existing.Status == ConnectionStatus.Connected))
                continue;

            ConnectDialogResult? dialog = null;
            if (!ConfigStore.HasUsableCredentials(request))
            {
                if (Dialogs is null) continue;
                var vm = new ConnectDialogViewModel(_config.Config.Groups);
                vm.LoadFrom(request, host.GroupId, host.DisplayName);
                dialog = await Dialogs.ShowConnectDialogAsync(vm);
                if (dialog is null) continue;
                request = dialog.Request;
            }
            StartConnection(request, dialog);
            started++;
        }
        return started;
    }

    [RelayCommand]
    private async Task ManageHosts()
    {
        if (Dialogs is null) return;
        await Dialogs.ShowHostsWindowAsync(new HostsViewModel(_config, Dialogs, hosts => ConnectSaved(hosts)));
    }

    [RelayCommand]
    private void RefreshAll()
    {
        var items = Connections.Connections.Where(c => !c.IsBusy && !c.IsDemo).ToList();
        if (items.Count == 0)
        {
            StatusText = Connections.Connections.Count == 0 ? "Nothing to refresh — connect to a host first."
                : Connections.Connections.All(c => c.IsDemo) ? "Demo data can't be refreshed — connect to a real host."
                : "Everything is already collecting.";
            return;
        }
        RefreshItems(items);
    }

    [RelayCommand]
    private void Refresh(ConnectionItem? item)
    {
        if (item is not null && !item.IsBusy && !item.IsDemo) RefreshItems([item]);
    }

    /// <summary>
    /// Re-collects sources live. Saved hosts with usable credentials start immediately (in parallel); saved hosts
    /// lacking them go through <see cref="ConnectSaved"/>, which prompts. Sources with no saved host (e.g. from a
    /// snapshot of another machine) open the connect dialog pre-filled. Not awaited, so commands stay enabled.
    /// </summary>
    private void RefreshItems(IReadOnlyList<ConnectionItem> items)
    {
        var needPrompt = new List<SavedHost>();
        var unknown = new List<ConnectionItem>();
        var started = 0;
        foreach (var item in items)
        {
            var saved = FindSavedHost(item);
            var request = RefreshRequest(item, saved);
            if (ConfigStore.HasUsableCredentials(request))
            {
                _ = Connections.ConnectAsync(request);
                started++;
            }
            else if (saved is not null)
            {
                needPrompt.Add(saved);
            }
            else
            {
                unknown.Add(item);
            }
        }
        StatusText = $"Refreshing {items.Count} source(s)…";
        Connections.Log("Refresh", $"Re-collecting {items.Count} source(s)");
        if (needPrompt.Count > 0) _ = ConnectSaved(needPrompt, includeConnected: true);
        if (unknown.Count > 0) _ = PromptAndConnect(unknown);
    }

    /// <summary>Asks for connection details for sources that have no saved host (typically loaded from a snapshot).</summary>
    private async Task PromptAndConnect(IReadOnlyList<ConnectionItem> items)
    {
        if (Dialogs is null) return;
        foreach (var item in items)
        {
            var vm = new ConnectDialogViewModel(_config.Config.Groups);
            vm.LoadFrom(item.Request with { Address = SplitAddress(item.Request.Address).Host, Port = SplitAddress(item.Request.Address).Port ?? item.Request.Port });
            if (await Dialogs.ShowConnectDialogAsync(vm) is { } result) StartConnection(result.Request, result);
        }
    }

    /// <summary>Snapshot sources may be keyed "host:port"; match them to saved hosts either way.</summary>
    private SavedHost? FindSavedHost(ConnectionItem item)
    {
        var (host, _) = SplitAddress(item.Request.Address);
        return _config.Config.FindHost(item.Request.Address, item.Platform) ?? _config.Config.FindHost(host, item.Platform);
    }

    private static (string Host, int? Port) SplitAddress(string address)
    {
        var colon = address.LastIndexOf(':');
        return colon > 0 && address.Count(c => c == ':') == 1 && int.TryParse(address[(colon + 1)..], out var port)
            ? (address[..colon], port)
            : (address, null);
    }

    private ConnectionRequest RefreshRequest(ConnectionItem item, SavedHost? saved)
    {
        // Prefer freshly stored credentials (the user may have edited them since).
        if (saved is not null)
        {
            var stored = _config.BuildRequest(saved);
            if (ConfigStore.HasUsableCredentials(stored)) return stored;
        }
        return item.Request;
    }

    [RelayCommand]
    private void Cancel(ConnectionItem? item)
    {
        if (item is not null) Connections.Cancel(item);
    }

    [RelayCommand]
    private void Disconnect(ConnectionItem? item)
    {
        if (item is not null) Connections.Remove(item);
    }

    [RelayCommand]
    private async Task ClearAll()
    {
        if (Dialogs is not null && HasConnections && !await Dialogs.ConfirmAsync("Disconnect all",
                $"Disconnect all {Connections.Connections.Count} source(s) and clear the inventory?\n\nSaved hosts and credentials are kept."))
            return;
        Connections.Clear();
        Details = null;
        StatusText = "Disconnected all sources.";
    }

    [RelayCommand]
    private void LoadDemo()
    {
        Connections.AddImported(SampleInventory.Create(), ConnectionManager.DemoOrigin);
    }

    [RelayCommand]
    private async Task CopyHint(ConnectionItem? item)
    {
        if (item is null || Dialogs is null) return;
        await Dialogs.SetClipboardTextAsync(string.IsNullOrEmpty(item.Hint) ? item.Message : $"{item.Message}\n\n{item.Hint}");
    }

    // ---------------- Files & exports ----------------

    private string Stamp => DateTime.Now.ToString("yyyyMMdd_HHmmss");

    private string? StartDir => _config.Config.Settings.LastExportDirectory;

    private void RememberDir(string path)
    {
        _config.Config.Settings.LastExportDirectory = Path.GetDirectoryName(path);
        _config.Save();
    }

    private async Task RunExport(string what, string path, Action<string> export)
    {
        try
        {
            IsBusy = true;
            StatusText = $"Exporting {what}…";
            await Task.Run(() => export(path));
            RememberDir(path);
            StatusText = $"Exported {what} to {path}";
            Connections.Log("Export", $"{what} → {path}");
        }
        catch (Exception ex)
        {
            StatusText = $"Export failed: {ex.Message}";
            Connections.Log("Export", $"{what} failed: {ex.Message}", isError: true);
            if (Dialogs is not null) await Dialogs.ShowMessageAsync("Export failed", ex.Message);
        }
        finally
        {
            IsBusy = Connections.AnyBusy;
        }
    }

    [RelayCommand(CanExecute = nameof(HasData))]
    private async Task ExportXlsx()
    {
        if (Dialogs is null) return;
        var path = await Dialogs.PickSaveFileAsync("Export RVTools-compatible workbook", $"HypervisorExplorer_{Stamp}.xlsx", "xlsx", StartDir);
        if (path is null) return;
        var inv = _inventory;
        await RunExport("RVTools workbook", path, p => InventoryExporter.ExportXlsx(inv, p));
    }

    [RelayCommand(CanExecute = nameof(HasData))]
    private async Task ExportCsv()
    {
        if (Dialogs is null || SelectedTable is not { } table) return;
        var path = await Dialogs.PickSaveFileAsync($"Export {table.Name} (current filter) to CSV", $"{table.Name}_{Stamp}.csv", "csv", StartDir);
        if (path is null) return;
        var rows = table.View.Cast<GridRow>().Select(r => r.Source).ToList();
        var columns = VisibleColumns().Select(c => c.Column).ToList();
        await RunExport($"{table.Name} CSV ({rows.Count} rows)", path, p => InventoryExporter.ExportCsv(table.Definition, rows, p, columns));
    }

    [RelayCommand(CanExecute = nameof(HasData))]
    private async Task ExportCsvZip()
    {
        if (Dialogs is null) return;
        var path = await Dialogs.PickSaveFileAsync("Export all tables as CSV (zip)", $"HypervisorExplorer_{Stamp}_csv.zip", "zip", StartDir);
        if (path is null) return;
        var inv = _inventory;
        await RunExport("all tables as CSV", path, p => InventoryExporter.ExportCsvZip(inv, p));
    }

    [RelayCommand(CanExecute = nameof(HasData))]
    private async Task ExportHtml()
    {
        if (Dialogs is null) return;
        var path = await Dialogs.PickSaveFileAsync("Export HTML estate report", $"HypervisorExplorer_Report_{Stamp}.html", "html", StartDir);
        if (path is null) return;
        var inv = _inventory;
        await RunExport("HTML report", path, p => HtmlReport.Write(inv, p));
    }

    [RelayCommand(CanExecute = nameof(HasData))]
    private async Task SaveInventory()
    {
        if (Dialogs is null) return;
        var path = await Dialogs.PickSaveFileAsync("Save inventory snapshot", $"HypervisorExplorer_{Stamp}.hvx.json", "json", StartDir);
        if (path is null) return;
        var snaps = _store.Snapshots;
        await RunExport("inventory snapshot", path, p => InventoryJson.Save(snaps, p));
    }

    [RelayCommand]
    private async Task SaveHyperVScript()
    {
        if (Dialogs is null) return;
        var path = await Dialogs.PickSaveFileAsync("Save Hyper-V collection script", "Collect-HyperV.ps1", "ps1", StartDir);
        if (path is null) return;
        await File.WriteAllTextAsync(path, HyperVCollector.GetCollectionScript(), new System.Text.UTF8Encoding(true));
        RememberDir(path);
        await Dialogs.ShowMessageAsync("Collection script saved",
            $"Saved to {path}.\n\nFor hosts you can't reach over WinRM, copy the script to the Hyper-V host and run it in an elevated PowerShell:\n\n" +
            "  .\\Collect-HyperV.ps1 -OutFile hyperv.json\n\nThen use Open → Import Hyper-V collection file… here.");
    }

    [RelayCommand]
    private async Task ImportHyperVFile()
    {
        if (Dialogs is null) return;
        var path = await Dialogs.PickOpenFileAsync("Import Hyper-V collection file", ["json"]);
        if (path is null) return;
        try
        {
            var name = Path.GetFileNameWithoutExtension(path);
            var request = new ConnectionRequest { Platform = Platform.HyperV, Address = name };
            var snap = await Task.Run(() => HyperVJsonMapper.MapFile(path, request));
            Connections.AddImported([snap], Path.GetFileName(path));
        }
        catch (Exception ex)
        {
            await Dialogs.ShowMessageAsync("Could not import file", ex.Message);
        }
    }

    [RelayCommand]
    private async Task OpenInventory()
    {
        if (Dialogs is null) return;
        var path = await Dialogs.PickOpenFileAsync("Open inventory snapshot", ["json"]);
        if (path is null) return;
        try
        {
            var snaps = await Task.Run(() => InventoryJson.Load(path));
            Connections.AddImported(snaps, Path.GetFileName(path));
        }
        catch (Exception ex)
        {
            await Dialogs.ShowMessageAsync("Could not open file", ex.Message);
        }
    }
}
