using System.Net;
using HypervisorExplorer.Core.Collection;

namespace HypervisorExplorer.Collectors.HyperV;

/// <summary>Turns WinRM / PowerShell failures into <see cref="CollectionException"/>s with actionable hints.</summary>
public static class HyperVErrorMapper
{
    public const string EnableRemotingHint =
        "WinRM is not enabled or a firewall is blocking it. On the target host run (as Administrator):\n  Enable-PSRemoting -Force";

    public static string TrustedHostsCommand(string address) =>
        $"Set-Item WSMan:\\localhost\\Client\\TrustedHosts -Value '{address}' -Concatenate -Force";

    public static string TrustedHostsHint(string address) =>
        "Kerberos cannot be used for this connection (IP address, or a computer outside your domain), so WinRM " +
        "requires the target to be in this computer's TrustedHosts list (or HTTPS). Use the host name if it is " +
        "domain-joined; otherwise supply credentials and run in an elevated PowerShell on this computer:\n  " +
        TrustedHostsCommand(address);

    public static string IpWithCurrentUserHint(string address) =>
        "Kerberos (current user) authentication does not work with an IP address. Either:\n" +
        "  1. Connect using the host name (e.g. HV01.contoso.local), or\n" +
        "  2. Supply explicit credentials and add the IP to WinRM TrustedHosts (elevated PowerShell on this computer):\n  " +
        TrustedHostsCommand(address);

    public const string AccessDeniedHint =
        "Check that:\n- The user name and password are correct (use DOMAIN\\user or user@domain)\n" +
        "- The account is in the Administrators (or Remote Management Users) group on the host\n" +
        "- The account is in Hyper-V Administrators so it can read Hyper-V configuration";

    public const string HyperVPermissionHint =
        "Add the account to the 'Hyper-V Administrators' (or local Administrators) group on the host. " +
        "When collecting this computer, run Hypervisor Explorer elevated (Run as administrator).";

    public const string HyperVMissingHint =
        "The host does not have the Hyper-V PowerShell module (is it a Hyper-V host?). On Windows Server run:\n" +
        "  Install-WindowsFeature Hyper-V-PowerShell";

    public const string WinRmUnreachableHint =
        "Make sure:\n- The host name or IP is correct and the host is powered on\n" +
        "- WinRM is enabled on the target (Enable-PSRemoting -Force)\n" +
        "- Firewalls allow WinRM (TCP 5985, or 5986 for HTTPS)";

    public const string CertificateHint =
        "The WinRM HTTPS certificate is not trusted or does not match the host name. Connect with the name in the " +
        "certificate, trust the issuing CA on this computer, or allow untrusted certificates for this host.";

    public const string NotWindowsHint =
        "Use 'Save collection script...' to get Collect-HyperV.ps1, run it on the Hyper-V host " +
        "(e.g. .\\Collect-HyperV.ps1 -OutFile C:\\Temp\\hyperv.json, elevated), then import the JSON file.";

    /// <summary>True when <paramref name="address"/> is a literal IPv4/IPv6 address.</summary>
    public static bool IsIpAddress(string address) =>
        IPAddress.TryParse(address.Trim().Trim('[', ']'), out _);

    /// <summary>Maps a failure envelope from Invoke-Collection.ps1 to a <see cref="CollectionException"/>.</summary>
    public static CollectionException FromEnvelope(string? stage, string? error, string? category, ConnectionRequest request)
    {
        var addr = request.Address;
        var msg = (error ?? "").Trim();
        if (msg.Length == 0) msg = "Unknown error";
        var text = msg + " " + category;

        bool Has(params string[] needles) => needles.Any(n => text.Contains(n, StringComparison.OrdinalIgnoreCase));

        if (Has("HYPERV_MODULE_MISSING", "The term 'Get-VM' is not recognized", "The term 'Get-VMHost' is not recognized"))
            return new CollectionException(CollectionFailure.Protocol, $"'{addr}' is not a Hyper-V host: {Strip(msg)}", HyperVMissingHint);

        if (Has("required permission", "authorization policy", "Hyper-V Administrators"))
            return new CollectionException(CollectionFailure.PermissionDenied, $"Permission denied reading Hyper-V on '{addr}': {Strip(msg)}", HyperVPermissionHint);

        if (Has("TrustedHosts"))
            return new CollectionException(CollectionFailure.AuthenticationFailed, $"WinRM authentication to '{addr}' failed: {msg}", TrustedHostsHint(addr));

        if (Has("certificate") && (request.UseSsl || Has("SSL", "HTTPS", "0x80338126")))
            return new CollectionException(CollectionFailure.Certificate, $"WinRM HTTPS certificate problem on '{addr}': {msg}", CertificateHint);

        if (Has("Access is denied", "AccessDenied", "logon failure", "LogonFailure", "user name or password", "0x80090311",
                "0x8009030e", "0x8009030c", "Kerberos", "credentials were rejected"))
            return new CollectionException(CollectionFailure.AuthenticationFailed, $"Access denied by '{addr}': {msg}", AccessDeniedHint);

        if (Has("WinRM cannot complete the operation", "0x80338012", "CannotConnect", "cannot find the computer",
                "network path was not found", "timed out", "No such host", "could not be resolved", "WinRM client cannot process the request",
                "The client cannot connect"))
            return new CollectionException(CollectionFailure.Unreachable, $"Cannot connect to '{addr}' over WinRM: {msg}", WinRmUnreachableHint);

        return stage switch
        {
            "auth" => new CollectionException(CollectionFailure.AuthenticationFailed, $"Authentication to '{addr}' failed: {msg}", AccessDeniedHint),
            "connect" => new CollectionException(CollectionFailure.Unreachable, $"Cannot connect to '{addr}': {msg}", WinRmUnreachableHint),
            _ => new CollectionException(CollectionFailure.Other, $"Hyper-V collection on '{addr}' failed: {Strip(msg)}"),
        };
    }

    /// <summary>Removes our internal "CODE: " prefixes from script error messages.</summary>
    private static string Strip(string msg)
    {
        foreach (var p in new[] { "HYPERV_MODULE_MISSING: ", "HYPERV_ACCESS: " })
        {
            var i = msg.IndexOf(p, StringComparison.Ordinal);
            if (i >= 0) return msg.Remove(i, p.Length);
        }
        return msg;
    }
}
