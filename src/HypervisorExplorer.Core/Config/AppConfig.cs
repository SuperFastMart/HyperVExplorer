using System.Text.Json;
using System.Text.Json.Serialization;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Core.Config;

/// <summary>A saved connection target.</summary>
public sealed class SavedHost
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Address { get; set; } = "";
    public Platform Platform { get; set; }
    public int? Port { get; set; }
    public string? DisplayName { get; set; }
    /// <summary>Group this host belongs to; the group's credentials apply when the host has none of its own.</summary>
    public string? GroupId { get; set; }
    /// <summary>Null = inherit from group (or prompt).</summary>
    public CredentialKind? CredentialKind { get; set; }
    public string? Username { get; set; }
    public string? ProtectedSecret { get; set; }
    public bool IgnoreCertificateErrors { get; set; } = true;
    public bool UseSsl { get; set; }
    public DateTimeOffset? LastConnected { get; set; }
    public string? LastError { get; set; }

    [JsonIgnore]
    public string Label => string.IsNullOrWhiteSpace(DisplayName) ? Address : $"{DisplayName} ({Address})";
}

/// <summary>A named set of hosts sharing credentials, e.g. a site or a cluster.</summary>
public sealed class HostGroup
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Name { get; set; } = "";
    public CredentialKind CredentialKind { get; set; } = CredentialKind.UsernamePassword;
    public string? Username { get; set; }
    public string? ProtectedSecret { get; set; }
}

public sealed class AppSettings
{
    public int MaxParallelConnections { get; set; } = 4;
    public int SnapshotAgeWarningDays { get; set; } = 7;
    public bool HideEmptyColumns { get; set; } = true;
    public bool SourcesExpanded { get; set; } = true;
    public bool CheckForUpdates { get; set; } = true;
    public string? Theme { get; set; } = "Dark";
    public string? LastExportDirectory { get; set; }
}

public sealed class AppConfig
{
    public int Version { get; set; } = 1;
    public List<SavedHost> Hosts { get; set; } = [];
    public List<HostGroup> Groups { get; set; } = [];
    public AppSettings Settings { get; set; } = new();

    public HostGroup? GroupOf(SavedHost host) =>
        host.GroupId is null ? null : Groups.FirstOrDefault(g => g.Id == host.GroupId);

    public SavedHost? FindHost(string address, Platform platform) =>
        Hosts.FirstOrDefault(h => h.Platform == platform && string.Equals(h.Address, address, StringComparison.OrdinalIgnoreCase));
}

/// <summary>Loads/saves <see cref="AppConfig"/> and resolves stored credentials into <see cref="ConnectionRequest"/>s.</summary>
public sealed class ConfigStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
    };

    private readonly object _gate = new();

    public ConfigStore(string? directory = null)
    {
        Directory = directory ?? DefaultDirectory();
        System.IO.Directory.CreateDirectory(Directory);
        Protector = new SecretProtector(Directory);
        Config = Load();
    }

    public string Directory { get; }
    public string ConfigPath => Path.Combine(Directory, "config.json");
    public ISecretProtector Protector { get; }
    public AppConfig Config { get; private set; }

    public static string DefaultDirectory()
    {
        var root = OperatingSystem.IsWindows()
            ? Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".config");
        return Path.Combine(root, "HypervisorExplorer");
    }

    private AppConfig Load()
    {
        if (!File.Exists(ConfigPath)) return new AppConfig();
        try
        {
            return JsonSerializer.Deserialize<AppConfig>(File.ReadAllText(ConfigPath), JsonOptions) ?? new AppConfig();
        }
        catch (JsonException)
        {
            // Keep the broken file for inspection rather than silently overwriting it.
            File.Copy(ConfigPath, ConfigPath + ".corrupt", overwrite: true);
            return new AppConfig();
        }
    }

    public void Save()
    {
        lock (_gate)
        {
            var temp = ConfigPath + ".tmp";
            File.WriteAllText(temp, JsonSerializer.Serialize(Config, JsonOptions));
            File.Move(temp, ConfigPath, overwrite: true);
        }
    }

    public void SetSecret(SavedHost host, string? secret) =>
        host.ProtectedSecret = string.IsNullOrEmpty(secret) ? null : Protector.Protect(secret);

    public void SetSecret(HostGroup group, string? secret) =>
        group.ProtectedSecret = string.IsNullOrEmpty(secret) ? null : Protector.Protect(secret);

    /// <summary>
    /// Builds a connection request using the host's own credentials, then its group's.
    /// Returns null secret when nothing usable is stored, so the caller can prompt.
    /// </summary>
    public ConnectionRequest BuildRequest(SavedHost host)
    {
        var group = Config.GroupOf(host);
        CredentialKind kind;
        string? user;
        string? secret;
        // A host set to integrated sign-in can't use it off Windows: fall back to its group's credentials.
        var hostKind = host.CredentialKind == CredentialKind.CurrentUser && !OperatingSystem.IsWindows() && group is not null
            ? null
            : host.CredentialKind;
        if (hostKind is { } hk)
        {
            kind = hk;
            user = host.Username;
            secret = host.ProtectedSecret is null ? null : Protector.Unprotect(host.ProtectedSecret);
        }
        else if (group is not null)
        {
            kind = group.CredentialKind;
            user = group.Username;
            secret = group.ProtectedSecret is null ? null : Protector.Unprotect(group.ProtectedSecret);
        }
        else
        {
            kind = host.Platform == Platform.HyperV && OperatingSystem.IsWindows() ? CredentialKind.CurrentUser : CredentialKind.UsernamePassword;
            user = null;
            secret = null;
        }

        return new ConnectionRequest
        {
            Platform = host.Platform,
            Address = host.Address,
            Port = host.Port,
            CredentialKind = kind,
            Username = user,
            Secret = secret,
            IgnoreCertificateErrors = host.IgnoreCertificateErrors,
            UseSsl = host.UseSsl,
            Group = group?.Name,
        };
    }

    /// <summary>True when the request has what it needs to connect without prompting.</summary>
    /// <remarks>Integrated (current user) sign-in only exists on Windows; elsewhere such hosts must prompt.</remarks>
    public static bool HasUsableCredentials(ConnectionRequest r) =>
        (r.CredentialKind == CredentialKind.CurrentUser && OperatingSystem.IsWindows())
        || (r.CredentialKind != CredentialKind.CurrentUser && !string.IsNullOrEmpty(r.Username) && !string.IsNullOrEmpty(r.Secret));

    /// <summary>
    /// Records a successful connection. Credentials identical to the host's group are not duplicated onto the host
    /// (so rotating the group password takes effect). A stored secret is kept only when <paramref name="saveSecret"/>
    /// is set and still belongs to the same account.
    /// </summary>
    public SavedHost Remember(ConnectionRequest request, bool saveSecret, string? groupId = null)
    {
        lock (_gate)
        {
            var host = Config.FindHost(request.Address, request.Platform);
            if (host is null)
            {
                host = new SavedHost { Address = request.Address, Platform = request.Platform };
                Config.Hosts.Add(host);
            }
            if (groupId is not null) host.GroupId = groupId;
            host.Port = request.Port;
            host.IgnoreCertificateErrors = request.IgnoreCertificateErrors;
            host.UseSsl = request.UseSsl;
            host.LastConnected = DateTimeOffset.Now;
            host.LastError = null;

            var group = Config.GroupOf(host);
            var sameAsGroup = group is not null && group.CredentialKind == request.CredentialKind
                && (request.CredentialKind == CredentialKind.CurrentUser
                    || (string.Equals(group.Username, request.Username, StringComparison.OrdinalIgnoreCase)
                        && group.ProtectedSecret is not null && Protector.Unprotect(group.ProtectedSecret) == request.Secret));
            if (sameAsGroup)
            {
                host.CredentialKind = null;
                host.Username = null;
                host.ProtectedSecret = null;
                return host;
            }

            var accountChanged = host.CredentialKind != request.CredentialKind
                || !string.Equals(host.Username, request.Username, StringComparison.OrdinalIgnoreCase);
            host.CredentialKind = request.CredentialKind;
            host.Username = request.Username;
            if (!saveSecret || request.CredentialKind == CredentialKind.CurrentUser) host.ProtectedSecret = null;
            else if (request.Secret is not null) SetSecret(host, request.Secret);
            else if (accountChanged) host.ProtectedSecret = null;
            return host;
        }
    }
}
