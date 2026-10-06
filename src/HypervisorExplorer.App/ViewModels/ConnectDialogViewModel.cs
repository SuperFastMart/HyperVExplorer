using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using HypervisorExplorer.Collectors;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Config;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.App.ViewModels;

public sealed record PlatformOption(Platform Platform, string Label);

public sealed record CredentialOption(CredentialKind Kind, string Label);

public sealed record GroupOption(string? Id, string Name);

/// <summary>What the connect / edit-host dialog returns.</summary>
public sealed record ConnectDialogResult(ConnectionRequest Request, bool SaveCredentials, string? GroupId, string? DisplayName);

public sealed partial class ConnectDialogViewModel : ObservableObject
{
    public static readonly IReadOnlyList<PlatformOption> PlatformOptions =
    [
        new(Platform.HyperV, "Microsoft Hyper-V"),
        new(Platform.Proxmox, "Proxmox VE"),
        new(Platform.VMware, "VMware vCenter / ESXi"),
    ];

    public ConnectDialogViewModel(IEnumerable<HostGroup> groups, bool editOnly = false)
    {
        EditOnly = editOnly;
        Groups = [new GroupOption(null, "(no group)"), .. groups.Select(g => new GroupOption(g.Id, g.Name))];
        _selectedGroup = Groups[0];
        _selectedPlatform = PlatformOptions[0];
        RefreshCredentialOptions();
    }

    /// <summary>When true the dialog saves the host without connecting.</summary>
    public bool EditOnly { get; }
    public string Title => TitleOverride ?? (EditOnly ? "Saved host" : "Connect");

    /// <summary>Optional dialog title, e.g. "Duplicate pve01".</summary>
    public string? TitleOverride { get; init; }

    /// <summary>Extra check run on OK (e.g. reject an address that is already saved); returns an error or null.</summary>
    public Func<ConnectionRequest, string?>? ExtraValidation { get; init; }
    public string PrimaryButtonText => EditOnly ? "Save" : "Connect";

    public ObservableCollection<CredentialOption> CredentialOptions { get; } = [];
    public IReadOnlyList<GroupOption> Groups { get; }

    [ObservableProperty] private PlatformOption _selectedPlatform;
    [ObservableProperty] private CredentialOption? _selectedCredential;
    [ObservableProperty] private GroupOption _selectedGroup;
    [ObservableProperty] private string _address = "";
    [ObservableProperty] private string _displayName = "";
    [ObservableProperty] private string _port = "";
    [ObservableProperty] private string _username = "";
    [ObservableProperty] private string _secret = "";
    [ObservableProperty] private bool _saveCredentials = true;
    [ObservableProperty] private bool _ignoreCertificateErrors = true;
    [ObservableProperty] private bool _useSsl;
    [ObservableProperty] private bool _expandCluster = true;
    [ObservableProperty] private string? _error;

    public bool IsHyperV => SelectedPlatform.Platform == Platform.HyperV;
    public bool ShowsCertificateOption => SelectedPlatform.Platform != Platform.HyperV;
    public bool NeedsCredentials => SelectedCredential?.Kind is CredentialKind.UsernamePassword or CredentialKind.ApiToken;
    public bool InheritsFromGroup => SelectedGroup.Id is not null;

    public string PortPlaceholder => CollectorRegistry.DefaultPort(SelectedPlatform.Platform, UseSsl).ToString();

    public string UsernameLabel => SelectedCredential?.Kind == CredentialKind.ApiToken ? "Token ID" : "Username";
    public string SecretLabel => SelectedCredential?.Kind == CredentialKind.ApiToken ? "Token secret" : "Password";

    public string UsernameWatermark => InheritsFromGroup
        ? $"leave blank to use the '{SelectedGroup.Name}' group's credentials"
        : (SelectedPlatform.Platform, SelectedCredential?.Kind) switch
    {
        (Platform.Proxmox, CredentialKind.ApiToken) => "user@realm!tokenname  e.g. root@pam!inventory",
        (Platform.Proxmox, _) => "root  (or user@pve / user@domain)",
        (Platform.VMware, _) => "administrator@vsphere.local or root",
        _ => @"DOMAIN\user or user@domain",
    };

    public string Hint => (SelectedPlatform.Platform, SelectedCredential?.Kind) switch
    {
        (Platform.HyperV, CredentialKind.CurrentUser) =>
            "Uses your Windows sign-in (Kerberos). Connect by hostname, not IP. Requires WinRM on the host (Enable-PSRemoting). " +
            "Failover Cluster nodes are discovered automatically.",
        (Platform.HyperV, _) =>
            "Requires WinRM on the host (Enable-PSRemoting). IP addresses must be in your WinRM TrustedHosts list. " +
            "Failover Cluster nodes are discovered automatically.",
        (Platform.Proxmox, CredentialKind.ApiToken) =>
            "For accounts with two-factor auth or unattended runs: create a token under Datacenter → Permissions → API Tokens " +
            "(PVEAuditor on / is enough). Connecting to one node collects the whole cluster.",
        (Platform.Proxmox, _) =>
            "Sign in with the same account you use for the Proxmox web UI (e.g. root, or user@pve). " +
            "Connecting to any one node collects the whole cluster.",
        (Platform.VMware, _) =>
            "Uses the vSphere API (HTTPS 443) like RVTools — no SSH needed. Point at vCenter for the whole estate, " +
            "or at an individual ESXi host. A read-only role is sufficient.",
        _ => "",
    };

    partial void OnSelectedPlatformChanged(PlatformOption value)
    {
        RefreshCredentialOptions();
        OnPropertyChanged(nameof(IsHyperV));
        OnPropertyChanged(nameof(ShowsCertificateOption));
        OnPropertyChanged(nameof(PortPlaceholder));
        OnPropertyChanged(nameof(UsernameWatermark));
        OnPropertyChanged(nameof(Hint));
    }

    partial void OnSelectedCredentialChanged(CredentialOption? value)
    {
        OnPropertyChanged(nameof(NeedsCredentials));
        OnPropertyChanged(nameof(UsernameLabel));
        OnPropertyChanged(nameof(SecretLabel));
        OnPropertyChanged(nameof(UsernameWatermark));
        OnPropertyChanged(nameof(Hint));
    }

    partial void OnSelectedGroupChanged(GroupOption value)
    {
        OnPropertyChanged(nameof(InheritsFromGroup));
        OnPropertyChanged(nameof(UsernameWatermark));
    }

    partial void OnUseSslChanged(bool value) => OnPropertyChanged(nameof(PortPlaceholder));

    private void RefreshCredentialOptions()
    {
        var previous = SelectedCredential?.Kind;
        CredentialOptions.Clear();
        switch (SelectedPlatform.Platform)
        {
            case Platform.HyperV:
                CredentialOptions.Add(new CredentialOption(CredentialKind.CurrentUser, "Current Windows user"));
                CredentialOptions.Add(new CredentialOption(CredentialKind.UsernamePassword, "Username and password"));
                break;
            case Platform.Proxmox:
                CredentialOptions.Add(new CredentialOption(CredentialKind.UsernamePassword, "Username and password"));
                CredentialOptions.Add(new CredentialOption(CredentialKind.ApiToken, "API token"));
                break;
            default:
                CredentialOptions.Add(new CredentialOption(CredentialKind.UsernamePassword, "Username and password"));
                break;
        }
        SelectedCredential = CredentialOptions.FirstOrDefault(o => o.Kind == previous) ?? CredentialOptions[0];
    }

    public void LoadFrom(ConnectionRequest request, string? groupId = null, string? displayName = null)
    {
        SelectedPlatform = PlatformOptions.First(p => p.Platform == request.Platform);
        Address = request.Address;
        Port = request.Port?.ToString() ?? "";
        SelectedCredential = CredentialOptions.FirstOrDefault(c => c.Kind == request.CredentialKind) ?? SelectedCredential;
        Username = request.Username ?? "";
        Secret = request.Secret ?? "";
        IgnoreCertificateErrors = request.IgnoreCertificateErrors;
        UseSsl = request.UseSsl;
        ExpandCluster = request.ExpandCluster;
        SelectedGroup = Groups.FirstOrDefault(g => g.Id == groupId) ?? Groups[0];
        DisplayName = displayName ?? "";
    }

    /// <summary>Validates input and builds the result, or sets <see cref="Error"/> and returns null.</summary>
    public ConnectDialogResult? TryBuild()
    {
        Error = null;
        var address = Address.Trim();
        if (address.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) address = address[8..];
        address = address.TrimEnd('/');
        int? port = null;
        var colon = address.LastIndexOf(':');
        if (colon > 0 && address.Count(c => c == ':') == 1 && int.TryParse(address[(colon + 1)..], out var embedded))
        {
            port = embedded;
            address = address[..colon];
        }
        if (address.Length == 0)
        {
            Error = "Enter a host name or IP address.";
            return null;
        }
        if (!string.IsNullOrWhiteSpace(Port))
        {
            if (!int.TryParse(Port, out var p) || p is < 1 or > 65535)
            {
                Error = "Port must be a number between 1 and 65535.";
                return null;
            }
            port = p;
        }

        var kind = SelectedCredential?.Kind ?? CredentialKind.UsernamePassword;
        var user = Username.Trim();
        if (NeedsCredentials && !InheritsFromGroup && !EditOnly)
        {
            if (user.Length == 0)
            {
                Error = $"Enter the {UsernameLabel.ToLowerInvariant()}.";
                return null;
            }
            if (Secret.Length == 0)
            {
                Error = $"Enter the {SecretLabel.ToLowerInvariant()}.";
                return null;
            }
        }
        if (kind == CredentialKind.ApiToken && user.Length > 0 && !user.Contains('!'))
        {
            Error = "Token ID must look like user@realm!tokenname.";
            return null;
        }

        var request = new ConnectionRequest
        {
            Platform = SelectedPlatform.Platform,
            Address = address,
            Port = port,
            CredentialKind = kind,
            Username = user.Length == 0 ? null : user,
            Secret = Secret.Length == 0 ? null : Secret,
            IgnoreCertificateErrors = IgnoreCertificateErrors,
            UseSsl = UseSsl,
            ExpandCluster = ExpandCluster,
            Group = SelectedGroup.Id is null ? null : SelectedGroup.Name,
        };
        if (ExtraValidation?.Invoke(request) is { } validationError)
        {
            Error = validationError;
            return null;
        }
        return new ConnectDialogResult(request, SaveCredentials, SelectedGroup.Id,
            string.IsNullOrWhiteSpace(DisplayName) ? null : DisplayName.Trim());
    }
}
