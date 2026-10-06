using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Core.Collection;

public enum CredentialKind
{
    /// <summary>Windows integrated auth (Kerberos/NTLM) as the signed-in user. Hyper-V only.</summary>
    CurrentUser,
    UsernamePassword,
    /// <summary>Proxmox API token: <see cref="ConnectionRequest.Username"/> holds "user@realm!tokenid".</summary>
    ApiToken,
}

/// <summary>Everything a collector needs to connect to one source.</summary>
public sealed record ConnectionRequest
{
    public required Platform Platform { get; init; }
    public required string Address { get; init; }
    /// <summary>Port override; null means the platform default (WinRM 5985/5986, PVE 8006, vSphere 443).</summary>
    public int? Port { get; init; }
    public CredentialKind CredentialKind { get; init; } = CredentialKind.UsernamePassword;
    public string? Username { get; init; }
    /// <summary>Password or token secret. Held in memory only for the duration of a collection.</summary>
    public string? Secret { get; init; }
    /// <summary>Accept self-signed / untrusted TLS certificates (default for PVE and ESXi, like RVTools).</summary>
    public bool IgnoreCertificateErrors { get; init; } = true;
    /// <summary>Hyper-V: connect to WinRM over HTTPS.</summary>
    public bool UseSsl { get; init; }
    /// <summary>Hyper-V: when the host is a Failover Cluster node, also collect every other node.</summary>
    public bool ExpandCluster { get; init; } = true;
    /// <summary>Site/group label from host groups; becomes the "Datacenter" in RVTools tables.</summary>
    public string? Group { get; init; }

    public string DisplayName => Port is null ? Address : $"{Address}:{Port}";

    public override string ToString() => $"{Platform} {DisplayName} ({CredentialKind}{(Username is null ? "" : " " + Username)})";
}

public enum CollectionFailure
{
    Unreachable,
    AuthenticationFailed,
    PermissionDenied,
    Certificate,
    Protocol,
    Cancelled,
    Other,
}

/// <summary>A collection failure with a category and an actionable hint for the user.</summary>
public sealed class CollectionException : Exception
{
    public CollectionException(CollectionFailure kind, string message, string? hint = null, Exception? inner = null)
        : base(message, inner)
    {
        Kind = kind;
        Hint = hint;
    }

    public CollectionFailure Kind { get; }
    public string? Hint { get; }
}

/// <summary>Collects inventory from one source. Implementations must be safe to call from a background thread.</summary>
public interface IInventoryCollector
{
    Platform Platform { get; }

    /// <summary>Connects, collects and returns a snapshot. Throws <see cref="CollectionException"/> on failure.</summary>
    Task<InventorySnapshot> CollectAsync(ConnectionRequest request, IProgress<string>? progress, CancellationToken cancellationToken);
}
