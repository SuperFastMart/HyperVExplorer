using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Config;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Tests.Core;

public class ConfigTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "hve-cfg-" + Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        if (Directory.Exists(_dir)) Directory.Delete(_dir, recursive: true);
    }

    [Fact]
    public void Secrets_round_trip_and_are_not_stored_in_plaintext()
    {
        var store = new ConfigStore(_dir);
        var host = new SavedHost { Address = "esx01", Platform = Platform.VMware, CredentialKind = CredentialKind.UsernamePassword, Username = "root" };
        store.SetSecret(host, "S3cret!");
        store.Config.Hosts.Add(host);
        store.Save();

        Assert.DoesNotContain("S3cret!", File.ReadAllText(store.ConfigPath));
        var reloaded = new ConfigStore(_dir);
        var req = reloaded.BuildRequest(reloaded.Config.Hosts.Single());
        Assert.Equal("root", req.Username);
        Assert.Equal("S3cret!", req.Secret);
        Assert.True(ConfigStore.HasUsableCredentials(req));
    }

    [Fact]
    public void Host_inherits_group_credentials_unless_overridden()
    {
        var store = new ConfigStore(_dir);
        var group = new HostGroup { Name = "London", CredentialKind = CredentialKind.UsernamePassword, Username = @"CORP\svc" };
        store.SetSecret(group, "grp");
        store.Config.Groups.Add(group);
        var inherits = new SavedHost { Address = "hv01", Platform = Platform.HyperV, GroupId = group.Id };
        var overrides = new SavedHost { Address = "hv02", Platform = Platform.HyperV, GroupId = group.Id, CredentialKind = CredentialKind.CurrentUser };
        store.Config.Hosts.AddRange([inherits, overrides]);

        var a = store.BuildRequest(inherits);
        Assert.Equal(@"CORP\svc", a.Username);
        Assert.Equal("grp", a.Secret);
        Assert.Equal("London", a.Group);

        var b = store.BuildRequest(overrides);
        Assert.Equal(CredentialKind.CurrentUser, b.CredentialKind);
        Assert.True(ConfigStore.HasUsableCredentials(b));
    }

    [Fact]
    public void Ungrouped_host_without_credentials_needs_prompt()
    {
        var store = new ConfigStore(_dir);
        var host = new SavedHost { Address = "pve1", Platform = Platform.Proxmox };
        Assert.False(ConfigStore.HasUsableCredentials(store.BuildRequest(host)));
    }

    [Fact]
    public void Remember_updates_existing_host()
    {
        var store = new ConfigStore(_dir);
        var req = new ConnectionRequest { Platform = Platform.Proxmox, Address = "pve1", CredentialKind = CredentialKind.ApiToken, Username = "root@pam!ro", Secret = "uuid" };
        store.Remember(req, saveSecret: true);
        store.Remember(req with { Port = 8007 }, saveSecret: false);
        var host = Assert.Single(store.Config.Hosts);
        Assert.Equal(8007, host.Port);
        Assert.Equal("uuid", store.BuildRequest(host).Secret);
    }

    [Fact]
    public void Imports_legacy_powershell_config()
    {
        var legacy = Path.Combine(_dir, "legacy.json");
        Directory.CreateDirectory(_dir);
        File.WriteAllText(legacy, """
            {
              "version": 3,
              "hosts": [
                { "address": "hv01.example.com", "type": "hyperv", "useCurrentUser": true, "lastConnected": "2025-01-01T10:00:00Z" },
                { "address": "192.0.2.50", "type": "proxmox", "useCurrentUser": false, "username": null }
              ],
              "groups": [
                { "name": "PVE Site", "type": "proxmox", "hosts": ["192.0.2.50", "192.0.2.51"], "pveAuthType": "token", "pveTokenId": "root@pam!audit" }
              ]
            }
            """);
        var store = new ConfigStore(_dir);
        var count = store.ImportLegacy(legacy);
        Assert.Equal(3, count);
        var group = Assert.Single(store.Config.Groups);
        Assert.Equal(CredentialKind.ApiToken, group.CredentialKind);
        Assert.Equal("root@pam!audit", group.Username);
        Assert.Equal(2, store.Config.Hosts.Count(h => h.GroupId == group.Id));
        Assert.Equal(CredentialKind.CurrentUser, store.Config.FindHost("hv01.example.com", Platform.HyperV)!.CredentialKind);
    }
}
