using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using HypervisorExplorer.Collectors;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.App.Services;

public enum ConnectionStatus
{
    Queued,
    Connecting,
    Connected,
    Failed,
    Cancelled,
    Imported,
}

/// <summary>One source in the connections panel. Mutated on the UI thread only.</summary>
public sealed partial class ConnectionItem : ObservableObject
{
    public ConnectionItem(ConnectionRequest request)
    {
        Request = request;
    }

    public ConnectionRequest Request { get; set; }
    public string Address => Request.Address;
    public Platform Platform => Request.Platform;
    public string PlatformLabel => Core.Tables.RvToolsTables.PlatformName(Platform);

    [ObservableProperty] private ConnectionStatus _status;
    [ObservableProperty] private string _message = "";
    [ObservableProperty] private string? _hint;
    [ObservableProperty] private int _vmCount;
    [ObservableProperty] private DateTimeOffset? _collectedAt;

    public bool IsBusy => Status is ConnectionStatus.Queued or ConnectionStatus.Connecting;

    partial void OnStatusChanged(ConnectionStatus value) => OnPropertyChanged(nameof(IsBusy));

    internal CancellationTokenSource? Cts { get; set; }
}

public sealed record ActivityEntry(DateTime Time, string Source, string Message, bool IsError)
{
    public string TimeText => Time.ToString("HH:mm:ss");
}

/// <summary>
/// Runs collections in the background with bounded parallelism and merges results into the
/// <see cref="InventoryStore"/>. All observable state is updated on the UI thread.
/// </summary>
public sealed class ConnectionManager
{
    private readonly InventoryStore _store;
    private readonly Func<Platform, IInventoryCollector> _collectorFactory;
    private SemaphoreSlim _gate;

    public ConnectionManager(InventoryStore store, int maxParallel = 4, Func<Platform, IInventoryCollector>? collectorFactory = null)
    {
        _store = store;
        _collectorFactory = collectorFactory ?? CollectorRegistry.Create;
        _gate = new SemaphoreSlim(Math.Max(1, maxParallel));
    }

    public ObservableCollection<ConnectionItem> Connections { get; } = [];
    public ObservableCollection<ActivityEntry> Activity { get; } = [];

    /// <summary>Raised on the UI thread after a collection succeeds; the bool says whether credentials should be remembered.</summary>
    public event Action<ConnectionRequest>? Succeeded;
    public event Action<ConnectionItem>? Failed;

    public bool AnyBusy => Connections.Any(c => c.IsBusy);

    public void SetParallelism(int max) => _gate = new SemaphoreSlim(Math.Max(1, max));

    public void Log(string source, string message, bool isError = false)
    {
        void Add()
        {
            Activity.Insert(0, new ActivityEntry(DateTime.Now, source, message, isError));
            while (Activity.Count > 1000) Activity.RemoveAt(Activity.Count - 1);
        }
        if (Dispatcher.UIThread.CheckAccess()) Add();
        else Dispatcher.UIThread.Post(Add);
    }

    public ConnectionItem? Find(string address) =>
        Connections.FirstOrDefault(c => string.Equals(c.Address, address, StringComparison.OrdinalIgnoreCase));

    /// <summary>Queues a collection. Re-running for an existing address refreshes it.</summary>
    public Task ConnectAsync(ConnectionRequest request)
    {
        var item = Find(request.Address);
        if (item is null)
        {
            item = new ConnectionItem(request);
            Connections.Add(item);
        }
        else
        {
            if (item.IsBusy) return Task.CompletedTask;
            item.Request = request;
        }

        item.Status = ConnectionStatus.Queued;
        item.Message = "Queued";
        item.Hint = null;
        item.Cts = new CancellationTokenSource();
        return RunAsync(item, item.Cts.Token);
    }

    private async Task RunAsync(ConnectionItem item, CancellationToken ct)
    {
        var request = item.Request;
        try
        {
            await _gate.WaitAsync(ct);
        }
        catch (OperationCanceledException)
        {
            item.Status = ConnectionStatus.Cancelled;
            item.Message = "Cancelled";
            return;
        }

        try
        {
            item.Status = ConnectionStatus.Connecting;
            item.Message = "Connecting…";
            Log(request.Address, $"Connecting ({item.PlatformLabel})");

            var progress = new Progress<string>(msg =>
            {
                item.Message = msg;
                Log(request.Address, msg);
            });

            var snapshot = await Task.Run(() => _collectorFactory(request.Platform).CollectAsync(request, progress, ct), ct);
            _store.Upsert(snapshot);

            item.Status = ConnectionStatus.Connected;
            item.VmCount = snapshot.VirtualMachines.Count;
            item.CollectedAt = DateTimeOffset.Now;
            item.Message = $"{snapshot.Hosts.Count} host(s), {snapshot.VirtualMachines.Count} VM(s)";
            foreach (var w in snapshot.Warnings) Log(request.Address, "Warning: " + w, isError: true);
            Log(request.Address, $"Collected {item.Message}");
            Succeeded?.Invoke(request);
        }
        catch (OperationCanceledException)
        {
            item.Status = ConnectionStatus.Cancelled;
            item.Message = "Cancelled";
            Log(request.Address, "Cancelled");
        }
        catch (CollectionException ex)
        {
            item.Status = ConnectionStatus.Failed;
            item.Message = ex.Message;
            item.Hint = ex.Hint;
            Log(request.Address, $"{ex.Kind}: {ex.Message}", isError: true);
            Failed?.Invoke(item);
        }
        catch (Exception ex)
        {
            item.Status = ConnectionStatus.Failed;
            item.Message = ex.Message;
            Log(request.Address, $"Error: {ex.Message}", isError: true);
            Failed?.Invoke(item);
        }
        finally
        {
            _gate.Release();
        }
    }

    public void Cancel(ConnectionItem item) => item.Cts?.Cancel();

    public void CancelAll()
    {
        foreach (var c in Connections) c.Cts?.Cancel();
    }

    public void Remove(ConnectionItem item)
    {
        item.Cts?.Cancel();
        Connections.Remove(item);
        _store.Remove(item.Address);
        Log(item.Address, "Disconnected");
    }

    public void Clear()
    {
        CancelAll();
        Connections.Clear();
        _store.Clear();
    }

    /// <summary>Adds snapshots loaded from a file (or demo data) as already-collected sources.</summary>
    public void AddImported(IEnumerable<InventorySnapshot> snapshots, string origin)
    {
        foreach (var snap in snapshots)
        {
            var request = new ConnectionRequest { Platform = snap.Source.Platform, Address = snap.Source.Address, Group = snap.Source.Group };
            var item = Find(request.Address) ?? new ConnectionItem(request);
            if (!Connections.Contains(item)) Connections.Add(item);
            item.Status = ConnectionStatus.Imported;
            item.VmCount = snap.VirtualMachines.Count;
            item.CollectedAt = snap.Source.CollectedAt;
            item.Message = $"{origin}: {snap.Hosts.Count} host(s), {snap.VirtualMachines.Count} VM(s)";
            _store.Upsert(snap);
            Log(request.Address, item.Message);
        }
    }
}
