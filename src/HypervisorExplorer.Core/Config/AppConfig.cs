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
        if (host.CredentialKind is { } hk)
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
            kind = host.Platform == Platform.HyperV ? CredentialKind.CurrentUser : CredentialKind.UsernamePassword;
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
    public static bool HasUsableCredentials(ConnectionRequest r) =>
        r.CredentialKind == CredentialKind.CurrentUser || (!string.IsNullOrEmpty(r.Username) && !string.IsNullOrEmpty(r.Secret));

    public SavedHost Remember(ConnectionRequest request, bool saveSecret)
    {
        lock (_gate)
        {
            var host = Config.FindHost(request.Address, request.Platform);
            if (host is null)
            {
                host = new SavedHost { Address = request.Address, Platform = request.Platform };
                Config.Hosts.Add(host);
            }
            host.Port = request.Port;
            host.IgnoreCertificateErrors = request.IgnoreCertificateErrors;
            host.UseSsl = request.UseSsl;
            host.LastConnected = DateTimeOffset.Now;
            host.LastError = null;
            var group = Config.GroupOf(host);
            var sameAsGroup = group is not null && group.CredentialKind == request.CredentialKind
                              && string.Equals(group.Username, request.Username, StringComparison.OrdinalIgnoreCase);
            if (!sameAsGroup)
            {
                host.CredentialKind = request.CredentialKind;
                host.Username = request.Username;
                if (saveSecret) SetSecret(host, request.Secret);
            }
            return host;
        }
    }

    /// <summary>
    /// Imports hosts and groups from the legacy PowerShell tool's %APPDATA%\HyperVExplorer\config.json.
    /// Secrets are re-encrypted when they can be decrypted (Windows, same user). Returns the number imported.
    /// </summary>
    public int ImportLegacy(string legacyPath)
    {
        using var doc = JsonDocument.Parse(File.ReadAllText(legacyPath));
        var root = doc.RootElement;
        var imported = 0;

        string? Str(JsonElement e, string name) =>
            e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;
        bool Bool(JsonElement e, string name) =>
            e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.True;
        Platform PlatformOf(string? type) => type is "proxmox" or "proxmox-pdm" ? Platform.Proxmox : Platform.HyperV;

        lock (_gate)
        {
            var groupIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (root.TryGetProperty("groups", out var groups) && groups.ValueKind == JsonValueKind.Array)
            {
                foreach (var g in groups.EnumerateArray())
                {
                    var name = Str(g, "name");
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    var platform = PlatformOf(Str(g, "type"));
                    var group = Config.Groups.FirstOrDefault(x => x.Name == name) ?? new HostGroup { Name = name };
                    if (!Config.Groups.Contains(group)) Config.Groups.Add(group);

                    if (platform == Platform.Proxmox)
                    {
                        var token = Str(g, "pveAuthType") == "token";
                        group.CredentialKind = token ? CredentialKind.ApiToken : CredentialKind.UsernamePassword;
                        group.Username = token ? Str(g, "pveTokenId") : Str(g, "pveUsername");
                        var legacySecret = Str(g, token ? "encryptedPveTokenSecret" : "encryptedPvePassword");
                        if (legacySecret is not null && SecretProtector.UnprotectLegacySecureString(legacySecret) is { } s)
                            SetSecret(group, s);
                    }
                    else
                    {
                        group.CredentialKind = Bool(g, "useCurrentUser") ? CredentialKind.CurrentUser : CredentialKind.UsernamePassword;
                        group.Username = Str(g, "username");
                        if (Str(g, "encryptedPassword") is { } enc && SecretProtector.UnprotectLegacySecureString(enc) is { } s)
                            SetSecret(group, s);
                    }

                    if (g.TryGetProperty("hosts", out var members) && members.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var m in members.EnumerateArray())
                        {
                            if (m.GetString() is not { Length: > 0 } addr) continue;
                            groupIds[addr] = group.Id;
                            var host = Config.FindHost(addr, platform);
                            if (host is null)
                            {
                                host = new SavedHost { Address = addr, Platform = platform };
                                Config.Hosts.Add(host);
                                imported++;
                            }
                            host.GroupId = group.Id;
                        }
                    }
                }
            }

            if (root.TryGetProperty("hosts", out var hosts) && hosts.ValueKind == JsonValueKind.Array)
            {
                foreach (var h in hosts.EnumerateArray())
                {
                    var addr = Str(h, "address");
                    if (string.IsNullOrWhiteSpace(addr)) continue;
                    var platform = PlatformOf(Str(h, "type"));
                    var host = Config.FindHost(addr, platform);
                    if (host is null)
                    {
                        host = new SavedHost { Address = addr, Platform = platform };
                        Config.Hosts.Add(host);
                        imported++;
                    }
                    if (groupIds.TryGetValue(addr, out var gid)) host.GroupId = gid;
                    if (DateTimeOffset.TryParse(Str(h, "lastConnected"), out var last)) host.LastConnected = last;
                    if (Bool(h, "useCurrentUser"))
                    {
                        host.CredentialKind = CredentialKind.CurrentUser;
                    }
                    else if (Str(h, "username") is { } user)
                    {
                        host.CredentialKind = CredentialKind.UsernamePassword;
                        host.Username = user;
                        if (Str(h, "encryptedPassword") is { } enc && SecretProtector.UnprotectLegacySecureString(enc) is { } s)
                            SetSecret(host, s);
                    }
                }
            }
        }
        Save();
        return imported;
    }

    public static string LegacyConfigPath() =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "HyperVExplorer", "config.json");
}
