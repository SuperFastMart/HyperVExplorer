using HypervisorExplorer.App.Services;
using HypervisorExplorer.App.ViewModels;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Config;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Tests.App;

public class HostsViewModelTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "hve-hosts-" + Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        if (Directory.Exists(_dir)) Directory.Delete(_dir, recursive: true);
    }

    /// <summary>Simulates the user editing the connect dialog: applies an edit, then presses OK.</summary>
    private sealed class FakeDialogs(Action<ConnectDialogViewModel> edit) : IDialogService
    {
        public string? LastError { get; private set; }

        public Task<ConnectDialogResult?> ShowConnectDialogAsync(ConnectDialogViewModel vm)
        {
            edit(vm);
            var result = vm.TryBuild();
            LastError = vm.Error;
            return Task.FromResult(result);
        }

        public Task<bool> ShowGroupDialogAsync(GroupEditViewModel vm) => Task.FromResult(false);
        public Task ShowHostsWindowAsync(HostsViewModel vm) => Task.CompletedTask;
        public Task<string?> PickSaveFileAsync(string title, string suggestedName, string extension, string? startDirectory = null) => Task.FromResult<string?>(null);
        public Task<string?> PickOpenFileAsync(string title, IReadOnlyList<string> extensions) => Task.FromResult<string?>(null);
        public Task<bool> ConfirmAsync(string title, string message) => Task.FromResult(true);
        public Task ShowMessageAsync(string title, string message) => Task.CompletedTask;
        public Task SetClipboardTextAsync(string text) => Task.CompletedTask;
    }

    private ConfigStore StoreWithHost(out SavedHost host, string? groupId = null)
    {
        var store = new ConfigStore(_dir);
        host = new SavedHost
        {
            Address = "10.44.102.41", Platform = Platform.Proxmox, Port = 8006, IgnoreCertificateErrors = true,
            CredentialKind = CredentialKind.UsernamePassword, Username = "root", GroupId = groupId,
        };
        store.SetSecret(host, "s3cret");
        store.Config.Hosts.Add(host);
        store.Save();
        return store;
    }

    [Fact]
    public async Task Duplicate_copies_settings_and_credentials_with_new_address()
    {
        var store = StoreWithHost(out var original);
        var dialogs = new FakeDialogs(vm => vm.Address = "10.44.102.42");
        var vm = new HostsViewModel(store, dialogs, _ => Task.CompletedTask);
        vm.SelectedHost = vm.Hosts.Single();

        await vm.DuplicateHostCommand.ExecuteAsync(null);

        Assert.Equal(2, store.Config.Hosts.Count);
        var copy = store.Config.Hosts.Single(h => h.Address == "10.44.102.42");
        Assert.Equal(Platform.Proxmox, copy.Platform);
        Assert.Equal(8006, copy.Port);
        var req = store.BuildRequest(copy);
        Assert.Equal("root", req.Username);
        Assert.Equal("s3cret", req.Secret);
        Assert.Equal("s3cret", store.BuildRequest(original).Secret); // original untouched
        Assert.Equal(copy, vm.SelectedHost!.Host);
    }

    [Fact]
    public async Task Duplicate_of_group_host_keeps_inheriting_from_group()
    {
        var store = new ConfigStore(_dir);
        var group = new HostGroup { Name = "UKDC1", Username = "root" };
        store.SetSecret(group, "grp");
        store.Config.Groups.Add(group);
        store.Config.Hosts.Add(new SavedHost { Address = "10.0.0.1", Platform = Platform.Proxmox, GroupId = group.Id });

        var vm = new HostsViewModel(store, new FakeDialogs(d => d.Address = "10.0.0.2"), _ => Task.CompletedTask);
        vm.SelectedHost = vm.Hosts.Single();
        await vm.DuplicateHostCommand.ExecuteAsync(null);

        var copy = store.Config.Hosts.Single(h => h.Address == "10.0.0.2");
        Assert.Equal(group.Id, copy.GroupId);
        Assert.Null(copy.CredentialKind);
        Assert.Equal("grp", store.BuildRequest(copy).Secret);
    }

    [Fact]
    public async Task Duplicate_with_unchanged_address_is_rejected()
    {
        var store = StoreWithHost(out _);
        var dialogs = new FakeDialogs(_ => { });
        var vm = new HostsViewModel(store, dialogs, _ => Task.CompletedTask);
        vm.SelectedHost = vm.Hosts.Single();

        await vm.DuplicateHostCommand.ExecuteAsync(null);

        Assert.Single(store.Config.Hosts);
        Assert.Contains("already saved", dialogs.LastError);
    }
}
