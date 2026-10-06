using System.Collections.ObjectModel;
using System.Collections.Specialized;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using HypervisorExplorer.App.Services;
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

        Connections = new ConnectionManager(_store, config.Config.Settings.MaxParallelConnections);
        Connections.Succeeded += OnConnectionSucceeded;
        Connections.Failed += item =>
        {
            if (_config.Config.FindHost(item.Address, item.Platform) is { } host)
            {
                host.LastError = item.Message;
                _config.Save();
            }
        };
        Connections.Connections.CollectionChanged += (_, _) => OnPropertyChanged(nameof(HasConnections));
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
        await StartConnection(result.Request, result);
    }

    private async Task StartConnection(ConnectionRequest request, ConnectDialogResult? fromDialog)
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
        await Connections.ConnectAsync(request);
    }

    private void OnConnectionSucceeded(ConnectionRequest request)
    {
        _pendingSaves.Remove(request.Address, out var dialog);
        var existing = _config.Config.FindHost(request.Address, request.Platform);
        var host = _config.Remember(request, dialog?.SaveCredentials ?? existing?.ProtectedSecret is not null);
        if (dialog is not null)
        {
            if (dialog.GroupId is not null) host.GroupId = dialog.GroupId;
            if (dialog.DisplayName is not null) host.DisplayName = dialog.DisplayName;
        }
        _config.Save();
    }

    /// <summary>Connects saved hosts, prompting only for those without usable credentials.</summary>
    public async Task ConnectSaved(IReadOnlyList<SavedHost> hosts)
    {
        foreach (var host in hosts)
        {
            var request = _config.BuildRequest(host);
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
            _ = StartConnection(request, dialog);
        }
    }

    [RelayCommand]
    private async Task ManageHosts()
    {
        if (Dialogs is null) return;
        await Dialogs.ShowHostsWindowAsync(new HostsViewModel(_config, Dialogs, ConnectSaved));
    }

    [RelayCommand]
    private async Task RefreshAll()
    {
        foreach (var item in Connections.Connections.Where(c => c.Status is ConnectionStatus.Connected or ConnectionStatus.Failed).ToList())
            await Connections.ConnectAsync(RefreshRequest(item));
    }

    [RelayCommand]
    private Task Refresh(ConnectionItem? item) =>
        item is null || item.Status == ConnectionStatus.Imported ? Task.CompletedTask : Connections.ConnectAsync(RefreshRequest(item));

    private ConnectionRequest RefreshRequest(ConnectionItem item)
    {
        // Prefer freshly stored credentials (the user may have edited them since).
        if (_config.Config.FindHost(item.Address, item.Platform) is { } saved)
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
        if (Dialogs is not null && HasData && !await Dialogs.ConfirmAsync("Clear inventory", "Disconnect all sources and clear the inventory?"))
            return;
        Connections.Clear();
        Details = null;
    }

    [RelayCommand]
    private void LoadDemo()
    {
        Connections.AddImported(SampleInventory.Create(), "Demo data");
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
