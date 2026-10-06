using System.Net.Sockets;
using System.Security.Authentication;
using System.Text.Json;
using HypervisorExplorer.Core.Collection;

namespace HypervisorExplorer.Collectors.Proxmox;

/// <summary>Maps transport/HTTP failures during connect to <see cref="CollectionException"/> with actionable hints.</summary>
internal static class PveErrors
{
    public const string PermissionHint =
        "Grant the PVEAuditor role on '/' (Datacenter → Permissions → Add, path '/', propagate). " +
        "For an API token with Privilege Separation enabled, grant the role to the token itself as well as the user.";

    public static CollectionException Map(Exception ex, ConnectionRequest request, Uri baseUri, CancellationToken ct)
    {
        if (ex is CollectionException ce) return ce;
        var target = $"{baseUri.Host}:{baseUri.Port}";
        var reachHint =
            $"Check that the Proxmox VE node is online, that port {baseUri.Port} is open (try https://{target} in a browser), " +
            "and that no firewall is blocking the connection. The PVE API listens on port 8006 by default.";

        switch (ex)
        {
            case OperationCanceledException when ct.IsCancellationRequested:
                return new CollectionException(CollectionFailure.Cancelled, $"Collection from {target} was cancelled.", inner: ex);

            case PveTfaRequiredException:
                return new CollectionException(CollectionFailure.AuthenticationFailed,
                    $"{request.Username} on {target} requires two-factor authentication.",
                    "Accounts with TFA cannot sign in non-interactively. Create an API token (Datacenter → Permissions → API Tokens) and connect with it.",
                    ex);

            case PveApiException { StatusCode: 401 } api:
                return new CollectionException(CollectionFailure.AuthenticationFailed,
                    $"Authentication rejected by {target}: {api.Reason}{(api.Detail is null ? "" : " — " + api.Detail)}",
                    AuthHint(request, target), ex);

            case PveApiException { StatusCode: 403 } api:
                return new CollectionException(CollectionFailure.PermissionDenied,
                    $"Permission denied on {target} ({api.Path}): {api.Reason}", PermissionHint, ex);

            case PveApiException api:
                return new CollectionException(CollectionFailure.Protocol,
                    $"{target} returned HTTP {api.StatusCode} {api.Reason} for {api.Path}{(api.Detail is null ? "" : ": " + api.Detail)}",
                    api.StatusCode == 404 || api.StatusCode == 501
                        ? $"The server at {target} does not look like a Proxmox VE API endpoint. Check the address and port (default 8006)."
                        : null, ex);

            case TimeoutException:
                return new CollectionException(CollectionFailure.Unreachable, $"Timed out talking to {target}.", reachHint, ex);

            case HttpRequestException hre:
                return MapHttp(hre, target, reachHint, request);

            case JsonException:
                return new CollectionException(CollectionFailure.Protocol,
                    $"{target} returned a response that is not Proxmox VE API JSON.",
                    "Check the address and port: the PVE API is served at https://<host>:8006/api2/json.", ex);

            default:
                return new CollectionException(CollectionFailure.Other, $"Collection from {target} failed: {ex.Message}", null, ex);
        }
    }

    private static CollectionException MapHttp(HttpRequestException hre, string target, string reachHint, ConnectionRequest request)
    {
        var socket = FindInner<SocketException>(hre);
        var tls = FindInner<AuthenticationException>(hre);
        if (tls is not null || hre.HttpRequestError == HttpRequestError.SecureConnectionError)
        {
            var hint = request.IgnoreCertificateErrors
                ? $"The TLS handshake failed. Check that {target} serves HTTPS (the PVE API uses https on port 8006)."
                : "The server certificate is not trusted. Enable 'Ignore certificate errors' (PVE uses a self-signed certificate by default), " +
                  "or trust the cluster CA (/etc/pve/pve-root-ca.pem) on this machine.";
            return new CollectionException(CollectionFailure.Certificate,
                $"TLS/certificate error connecting to {target}: {(tls ?? (Exception)hre).Message}", hint, hre);
        }

        if (hre.HttpRequestError == HttpRequestError.NameResolutionError || socket?.SocketErrorCode == SocketError.HostNotFound)
            return new CollectionException(CollectionFailure.Unreachable, $"Cannot resolve host name '{target}'.",
                "Check the spelling of the host name, or use the node's IP address.", hre);

        if (socket is not null || hre.HttpRequestError is HttpRequestError.ConnectionError)
        {
            var what = socket?.SocketErrorCode switch
            {
                SocketError.ConnectionRefused => "connection refused",
                SocketError.TimedOut => "connection timed out",
                SocketError.HostUnreachable or SocketError.NetworkUnreachable => "host unreachable",
                _ => socket?.Message ?? hre.Message,
            };
            return new CollectionException(CollectionFailure.Unreachable, $"Cannot reach {target}: {what}.", reachHint, hre);
        }

        if (hre.HttpRequestError is HttpRequestError.InvalidResponse or HttpRequestError.ResponseEnded)
            return new CollectionException(CollectionFailure.Protocol, $"{target} sent an invalid HTTP response: {hre.Message}",
                "Check the port: the PVE API is HTTPS on 8006.", hre);

        return new CollectionException(CollectionFailure.Unreachable, $"Cannot reach {target}: {hre.Message}", reachHint, hre);
    }

    private static string AuthHint(ConnectionRequest request, string target)
    {
        if (request.CredentialKind == CredentialKind.ApiToken)
        {
            var (tokenId, _) = PveApiClient.NormalizeToken(request.Username, request.Secret);
            var formatNote = tokenId.Contains('!') && tokenId.Contains('@')
                ? ""
                : $"  - '{tokenId}' is not in the form user@realm!tokenname\n";
            return "For API token auth, verify:\n" + formatNote +
                   $"  - Token ID: {tokenId} (format: user@realm!tokenname, e.g. root@pam!explorer)\n" +
                   "  - Secret is the full UUID shown once when the token was created\n" +
                   "  - The token has not been revoked or expired\n" +
                   "  - The token has PVEAuditor (or higher) on /\n" +
                   $"Test with curl:\n  curl -k -H 'Authorization: PVEAPIToken={tokenId}=SECRET' https://{target}/api2/json/version";
        }

        return "For username/password auth:\n" +
               "  - Username format: user@realm (e.g. root@pam or admin@pve); a name without @ is treated as @pam\n" +
               "  - Check the password is correct\n" +
               "  - Accounts with two-factor authentication must use an API token instead";
    }

    private static T? FindInner<T>(Exception ex) where T : Exception
    {
        for (var e = ex; e is not null; e = e.InnerException)
            if (e is T t) return t;
        return null;
    }
}
