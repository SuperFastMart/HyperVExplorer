using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using HypervisorExplorer.App.Services;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Config;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.App.ViewModels;

public sealed partial class SavedHostItem : ObservableObject
{
    public SavedHostItem(SavedHost host, AppConfig config)
    {
        Host = host;
        Refresh(config);
    }

    public SavedHost Host { get; }

    [ObservableProperty] private bool _isChecked;
    [ObservableProperty] private string _label = "";
    [ObservableProperty] private string _platform = "";
    [ObservableProperty] private string _group = "";
    [ObservableProperty] private string _credentials = "";
    [ObservableProperty] private string _lastConnected = "";

    public void Refresh(AppConfig config)
    {
        Label = Host.Label;
        Platform = RvToolsTables.PlatformName(Host.Platform);
        var group = config.GroupOf(Host);
        Group = group?.Name ?? "";
        Credentials = Host.CredentialKind switch
        {
            CredentialKind.CurrentUser => "Current user",
            CredentialKind.ApiToken => $"Token {Host.Username}",
            CredentialKind.UsernamePassword => Host.Username + (Host.ProtectedSecret is null ? " (no password)" : ""),
            null when group is not null => $"From group: {group.Username ?? group.CredentialKind.ToString()}",
            _ => "Prompt",
        };
        LastConnected = Host.LastConnected?.LocalDateTime.ToString("yyyy-MM-dd HH:mm") ?? "never";
    }
}

public sealed partial class GroupEditViewModel : ObservableObject
{
    public static readonly IReadOnlyList<CredentialOption> KindOptions =
    [
        new(CredentialKind.UsernamePassword, "Username and password"),
        new(CredentialKind.ApiToken, "Proxmox API token"),
        new(CredentialKind.CurrentUser, "Current Windows user (Hyper-V)"),
    ];

    public GroupEditViewModel(HostGroup? group)
    {
        _name = group?.Name ?? "";
        _selectedKind = KindOptions.First(k => k.Kind == (group?.CredentialKind ?? CredentialKind.UsernamePassword));
        _username = group?.Username ?? "";
        HasSavedSecret = group?.ProtectedSecret is not null;
    }

    [ObservableProperty] private string _name;
    [ObservableProperty] private CredentialOption _selectedKind;
    [ObservableProperty] private string _username;
    [ObservableProperty] private string _secret = "";
    [ObservableProperty] private string? _error;

    public bool HasSavedSecret { get; }
    public string SecretWatermark => HasSavedSecret ? "(unchanged)" : "";
    public bool NeedsCredentials => SelectedKind.Kind != CredentialKind.CurrentUser;

    partial void OnSelectedKindChanged(CredentialOption value) => OnPropertyChanged(nameof(NeedsCredentials));

    public bool Validate()
    {
        Error = string.IsNullOrWhiteSpace(Name) ? "Enter a group name." : null;
        return Error is null;
    }
}

/// <summary>Saved hosts and groups: edit credentials once, connect many.</summary>
public sealed partial class HostsViewModel : ObservableObject
{
    private readonly ConfigStore _config;
    private readonly IDialogService _dialogs;
    private readonly Func<IReadOnlyList<SavedHost>, Task<int>> _connect;

    public HostsViewModel(ConfigStore config, IDialogService dialogs, Func<IReadOnlyList<SavedHost>, Task<int>> connect)
    {
        _config = config;
        _dialogs = dialogs;
        _connect = connect;
        Reload();
    }

    public ObservableCollection<SavedHostItem> Hosts { get; } = [];
    public ObservableCollection<HostGroup> Groups { get; } = [];

    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(EditHostCommand), nameof(DuplicateHostCommand), nameof(DeleteHostCommand))]
    private SavedHostItem? _selectedHost;

    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(EditGroupCommand), nameof(DeleteGroupCommand), nameof(ConnectGroupCommand))]
    private HostGroup? _selectedGroup;

    [ObservableProperty] private string? _status;

    private void Reload()
    {
        Hosts.Clear();
        foreach (var h in _config.Config.Hosts.OrderBy(h => _config.Config.GroupOf(h)?.Name).ThenBy(h => h.Address))
            Hosts.Add(new SavedHostItem(h, _config.Config));
        Groups.Clear();
        foreach (var g in _config.Config.Groups.OrderBy(g => g.Name)) Groups.Add(g);
    }

    [RelayCommand]
    private async Task AddHost()
    {
        var vm = new ConnectDialogViewModel(_config.Config.Groups, editOnly: true) { ExtraValidation = r => RejectExisting(r, null) };
        if (await _dialogs.ShowConnectDialogAsync(vm) is not { } result) return;
        SaveNewHost(result);
    }

    /// <summary>Copies the selected host's settings and credentials into a new host; only the address needs changing.</summary>
    [RelayCommand(CanExecute = nameof(HasHost))]
    private async Task DuplicateHost()
    {
        if (SelectedHost is not { } item) return;
        var source = item.Host;
        var vm = new ConnectDialogViewModel(_config.Config.Groups, editOnly: true)
        {
            TitleOverride = $"Duplicate {source.Label}",
            ExtraValidation = r => RejectExisting(r, null),
        };
        var request = source.CredentialKind is null
            ? new ConnectionRequest { Platform = source.Platform, Address = source.Address, Port = source.Port, IgnoreCertificateErrors = source.IgnoreCertificateErrors, UseSsl = source.UseSsl }
            : _config.BuildRequest(source);
        vm.LoadFrom(request, source.GroupId);
        vm.SaveCredentials = source.ProtectedSecret is not null || source.CredentialKind is null;
        if (await _dialogs.ShowConnectDialogAsync(vm) is not { } result) return;
        SaveNewHost(result);
    }

    private void SaveNewHost(ConnectDialogResult result)
    {
        var r = result.Request;
        var host = new SavedHost
        {
            Address = r.Address,
            Platform = r.Platform,
            Port = r.Port,
            IgnoreCertificateErrors = r.IgnoreCertificateErrors,
            UseSsl = r.UseSsl,
            GroupId = result.GroupId,
            DisplayName = result.DisplayName,
        };
        if (result.GroupId is not null && r.Username is null && r.CredentialKind != CredentialKind.CurrentUser)
        {
            host.CredentialKind = null; // inherit the group's credentials
        }
        else
        {
            host.CredentialKind = r.CredentialKind;
            host.Username = r.Username;
            if (result.SaveCredentials) _config.SetSecret(host, r.Secret);
        }
        _config.Config.Hosts.Add(host);
        _config.Save();
        Reload();
        SelectedHost = Hosts.FirstOrDefault(h => h.Host == host);
        Status = $"Saved {host.Label}.";
    }

    private static string Summarise(int started, int requested)
    {
        var skipped = requested - started;
        var text = $"Queued {started} host(s).";
        return skipped > 0 ? text + $" {skipped} already connected or skipped — use ⟳ Refresh to re-collect." : text;
    }

    private string? RejectExisting(ConnectionRequest r, SavedHost? except) =>
        _config.Config.Hosts.Any(h => h != except && h.Platform == r.Platform
                                      && string.Equals(h.Address, r.Address, StringComparison.OrdinalIgnoreCase))
            ? $"{r.Address} is already saved as a {RvToolsTables.PlatformName(r.Platform)} host. Change the address."
            : null;

    private bool HasHost => SelectedHost is not null;

    [RelayCommand(CanExecute = nameof(HasHost))]
    private async Task EditHost()
    {
        if (SelectedHost is not { } item) return;
        var host = item.Host;
        var vm = new ConnectDialogViewModel(_config.Config.Groups, editOnly: true) { ExtraValidation = r => RejectExisting(r, host) };
        var request = host.CredentialKind is null
            ? new ConnectionRequest { Platform = host.Platform, Address = host.Address, Port = host.Port, IgnoreCertificateErrors = host.IgnoreCertificateErrors, UseSsl = host.UseSsl }
            : _config.BuildRequest(host);
        vm.LoadFrom(request, host.GroupId, host.DisplayName);
        if (await _dialogs.ShowConnectDialogAsync(vm) is not { } result) return;

        var r = result.Request;
        host.Address = r.Address;
        host.Platform = r.Platform;
        host.Port = r.Port;
        host.IgnoreCertificateErrors = r.IgnoreCertificateErrors;
        host.UseSsl = r.UseSsl;
        host.GroupId = result.GroupId;
        host.DisplayName = result.DisplayName;
        if (result.GroupId is not null && r.Username is null && r.CredentialKind != CredentialKind.CurrentUser)
        {
            host.CredentialKind = null;
            host.Username = null;
            host.ProtectedSecret = null;
        }
        else
        {
            host.CredentialKind = r.CredentialKind;
            host.Username = r.Username;
            if (result.SaveCredentials) _config.SetSecret(host, r.Secret);
            else host.ProtectedSecret = null;
        }
        _config.Save();
        Reload();
    }

    [RelayCommand(CanExecute = nameof(HasHost))]
    private async Task DeleteHost()
    {
        var checkedHosts = Hosts.Where(h => h.IsChecked).Select(h => h.Host).ToList();
        if (checkedHosts.Count == 0 && SelectedHost is not null) checkedHosts.Add(SelectedHost.Host);
        if (checkedHosts.Count == 0) return;
        if (!await _dialogs.ConfirmAsync("Remove saved hosts", $"Remove {checkedHosts.Count} saved host(s) and their stored credentials?")) return;
        foreach (var h in checkedHosts) _config.Config.Hosts.Remove(h);
        _config.Save();
        Reload();
    }

    [RelayCommand]
    private async Task ConnectChecked()
    {
        var selected = Hosts.Where(h => h.IsChecked).Select(h => h.Host).ToList();
        if (selected.Count == 0 && SelectedHost is not null) selected.Add(SelectedHost.Host);
        if (selected.Count == 0)
        {
            Status = "Tick the hosts to connect, or select one.";
            return;
        }
        Status = Summarise(await _connect(selected), selected.Count);
    }

    [RelayCommand]
    private void SelectAll()
    {
        var target = Hosts.Any(h => !h.IsChecked);
        foreach (var h in Hosts) h.IsChecked = target;
    }

    [RelayCommand]
    private async Task AddGroup()
    {
        var vm = new GroupEditViewModel(null);
        if (!await _dialogs.ShowGroupDialogAsync(vm)) return;
        var group = new HostGroup();
        Apply(vm, group);
        _config.Config.Groups.Add(group);
        _config.Save();
        Reload();
    }

    private bool HasGroup => SelectedGroup is not null;

    [RelayCommand(CanExecute = nameof(HasGroup))]
    private async Task EditGroup()
    {
        if (SelectedGroup is not { } group) return;
        var vm = new GroupEditViewModel(group);
        if (!await _dialogs.ShowGroupDialogAsync(vm)) return;
        Apply(vm, group);
        _config.Save();
        Reload();
    }

    private void Apply(GroupEditViewModel vm, HostGroup group)
    {
        group.Name = vm.Name.Trim();
        group.CredentialKind = vm.SelectedKind.Kind;
        group.Username = vm.NeedsCredentials && vm.Username.Length > 0 ? vm.Username.Trim() : null;
        if (!vm.NeedsCredentials) group.ProtectedSecret = null;
        else if (vm.Secret.Length > 0) _config.SetSecret(group, vm.Secret);
    }

    [RelayCommand(CanExecute = nameof(HasGroup))]
    private async Task DeleteGroup()
    {
        if (SelectedGroup is not { } group) return;
        if (!await _dialogs.ConfirmAsync("Delete group", $"Delete group '{group.Name}'? Hosts are kept but lose the group's credentials.")) return;
        foreach (var h in _config.Config.Hosts.Where(h => h.GroupId == group.Id)) h.GroupId = null;
        _config.Config.Groups.Remove(group);
        _config.Save();
        Reload();
    }

    [RelayCommand(CanExecute = nameof(HasGroup))]
    private async Task ConnectGroup()
    {
        if (SelectedGroup is not { } group) return;
        var members = _config.Config.Hosts.Where(h => h.GroupId == group.Id).ToList();
        if (members.Count == 0)
        {
            Status = $"Group '{group.Name}' has no hosts. Edit a host to assign it.";
            return;
        }
        Status = Summarise(await _connect(members), members.Count) + $" (group '{group.Name}')";
    }
}
