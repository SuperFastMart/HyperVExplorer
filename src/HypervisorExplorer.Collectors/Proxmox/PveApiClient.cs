using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using HypervisorExplorer.Core.Collection;

namespace HypervisorExplorer.Collectors.Proxmox;

/// <summary>A non-success HTTP status from the PVE API. PVE puts its error text in the HTTP reason phrase.</summary>
internal sealed class PveApiException : Exception
{
    public PveApiException(int statusCode, string? reason, string path, string? detail = null)
        : base($"HTTP {statusCode} {reason} ({path}){(string.IsNullOrWhiteSpace(detail) ? "" : ": " + detail)}")
    {
        StatusCode = statusCode;
        Reason = reason;
        Path = path;
        Detail = detail;
    }

    public int StatusCode { get; }
    public string? Reason { get; }
    public string Path { get; }
    public string? Detail { get; }
}

/// <summary>Raised when the PVE ticket endpoint demands a second factor.</summary>
internal sealed class PveTfaRequiredException() : Exception("The account requires two-factor authentication.");

/// <summary>
/// Thin PVE REST client over <see cref="HttpClient"/>: authentication (API token or ticket), GET with a per-request
/// timeout, and a concurrency gate shared by every request so callers can fan out freely.
/// </summary>
internal sealed class PveApiClient : IDisposable
{
    public const int DefaultPort = 8006;

    private readonly HttpClient _http;
    private readonly SemaphoreSlim _gate;
    private readonly TimeSpan _defaultTimeout;
    private string? _tokenHeader;
    private string? _ticket;
    private string? _csrfToken;

    public PveApiClient(HttpMessageHandler handler, bool disposeHandler, Uri baseUri, int maxConcurrency, TimeSpan defaultTimeout)
    {
        _http = new HttpClient(handler, disposeHandler) { Timeout = Timeout.InfiniteTimeSpan };
        _gate = new SemaphoreSlim(Math.Max(1, maxConcurrency));
        _defaultTimeout = defaultTimeout;
        BaseUri = baseUri;
    }

    /// <summary>e.g. https://pve1:8006/api2/json</summary>
    public Uri BaseUri { get; }

    /// <summary>Builds the API base URI from a user-entered address ("pve1", "10.0.0.5", "https://pve1:8006/", "fe80::1").</summary>
    public static Uri BuildBaseUri(string address, int? port)
    {
        var a = address.Trim().TrimEnd('/');
        string host;
        int? embeddedPort = null;
        if (a.Contains("://", StringComparison.Ordinal) && Uri.TryCreate(a, UriKind.Absolute, out var u))
        {
            host = u.Host;
            if (!u.IsDefaultPort) embeddedPort = u.Port;
        }
        else if (a.StartsWith('[') && a.Contains(']'))
        {
            var end = a.IndexOf(']');
            host = a[1..end];
            if (end + 1 < a.Length && a[end + 1] == ':' && int.TryParse(a[(end + 2)..], out var p)) embeddedPort = p;
        }
        else if (a.Count(c => c == ':') == 1)
        {
            var parts = a.Split(':');
            host = parts[0];
            if (int.TryParse(parts[1], out var p)) embeddedPort = p;
        }
        else
        {
            host = a; // hostname, IPv4, or bare IPv6
        }

        var builder = new UriBuilder(Uri.UriSchemeHttps, host.Trim('[', ']'), port ?? embeddedPort ?? DefaultPort, "/api2/json");
        return builder.Uri;
    }

    /// <summary>Escapes a path segment (node, storage id).</summary>
    public static string Seg(string s) => Uri.EscapeDataString(s);

    /// <summary>
    /// Prepares credentials. Tokens are sent on every request; passwords are exchanged for a ticket
    /// (POST /access/ticket) which is then sent as the PVEAuthCookie.
    /// </summary>
    public async Task LoginAsync(ConnectionRequest request, CancellationToken ct)
    {
        switch (request.CredentialKind)
        {
            case CredentialKind.ApiToken:
            {
                var (tokenId, secret) = NormalizeToken(request.Username, request.Secret);
                _tokenHeader = $"PVEAPIToken={tokenId}={secret}";
                break;
            }
            case CredentialKind.UsernamePassword:
            {
                var user = NormalizeUser(request.Username);
                using var form = new FormUrlEncodedContent(
                [
                    new("username", user),
                    new("password", request.Secret ?? ""),
                ]);
                var data = await SendAsync(HttpMethod.Post, "/access/ticket", form, null, ct).ConfigureAwait(false);
                if (data.Bool("NeedTFA") == true) throw new PveTfaRequiredException();
                _ticket = data.Str("ticket");
                _csrfToken = data.Str("CSRFPreventionToken");
                if (string.IsNullOrEmpty(_ticket))
                    throw new PveApiException(401, "authentication failure", "/access/ticket", "no ticket returned");
                break;
            }
            default:
                throw new CollectionException(CollectionFailure.AuthenticationFailed,
                    "Proxmox VE does not support Windows integrated authentication.",
                    "Use an API token (user@realm!tokenid + secret) or a username@realm and password.");
        }
    }

    /// <summary>"root" → "root@pam"; realm-qualified names are kept.</summary>
    public static string NormalizeUser(string? username)
    {
        var u = (username ?? "").Trim();
        return u.Length == 0 || u.Contains('@') ? u : u + "@pam";
    }

    /// <summary>
    /// Accepts "user@realm!tokenid" + secret, but also tolerates a pasted "PVEAPIToken=user@realm!id=secret"
    /// or "user@realm!id=secret" in the username field.
    /// </summary>
    public static (string TokenId, string Secret) NormalizeToken(string? username, string? secret)
    {
        var id = (username ?? "").Trim();
        var sec = (secret ?? "").Trim();
        if (id.StartsWith("PVEAPIToken=", StringComparison.OrdinalIgnoreCase)) id = id["PVEAPIToken=".Length..];
        var eq = id.IndexOf('=');
        if (eq > 0)
        {
            if (sec.Length == 0) sec = id[(eq + 1)..];
            id = id[..eq];
        }
        var bang = id.IndexOf('!');
        if (bang > 0 && !id[..bang].Contains('@')) id = id[..bang] + "@pam" + id[bang..];
        return (id, sec);
    }

    /// <summary>GET a path under /api2/json and return its "data" element. Throws on any failure.</summary>
    public Task<JsonElement> GetAsync(string path, CancellationToken ct, TimeSpan? timeout = null) =>
        SendAsync(HttpMethod.Get, path, null, timeout, ct);

    /// <summary>GET for optional data: returns null on HTTP errors and timeouts; still honours cancellation.</summary>
    public async Task<JsonElement?> TryGetAsync(string path, CancellationToken ct, TimeSpan? timeout = null)
    {
        try
        {
            return await GetAsync(path, ct, timeout).ConfigureAwait(false);
        }
        catch (Exception ex) when (IsSoftFailure(ex, ct))
        {
            return null;
        }
    }

    /// <summary>True for failures of one optional call (HTTP status, timeout, bad JSON) but not user cancellation.</summary>
    public static bool IsSoftFailure(Exception ex, CancellationToken ct) =>
        !ct.IsCancellationRequested && ex is PveApiException or TimeoutException or JsonException or HttpRequestException;

    private async Task<JsonElement> SendAsync(HttpMethod method, string path, HttpContent? content, TimeSpan? timeout, CancellationToken ct)
    {
        using var req = new HttpRequestMessage(method, BaseUri + path) { Content = content };
        req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        if (_tokenHeader is not null) req.Headers.TryAddWithoutValidation("Authorization", _tokenHeader);
        if (_ticket is not null)
        {
            req.Headers.TryAddWithoutValidation("Cookie", "PVEAuthCookie=" + _ticket);
            if (method != HttpMethod.Get && _csrfToken is not null)
                req.Headers.TryAddWithoutValidation("CSRFPreventionToken", _csrfToken);
        }

        var limit = timeout ?? _defaultTimeout;
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        await _gate.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            cts.CancelAfter(limit);
            using var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseContentRead, cts.Token).ConfigureAwait(false);
            var body = await resp.Content.ReadAsStringAsync(cts.Token).ConfigureAwait(false);
            if (!resp.IsSuccessStatusCode)
                throw new PveApiException((int)resp.StatusCode, resp.ReasonPhrase, path, ErrorDetail(body));
            if (string.IsNullOrWhiteSpace(body)) return default;
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("data", out var data) ? data.Clone() : default;
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            throw new TimeoutException($"Request timed out after {limit.TotalSeconds:0}s ({path})");
        }
        finally
        {
            _gate.Release();
        }
    }

    private static string? ErrorDetail(string body)
    {
        if (string.IsNullOrWhiteSpace(body)) return null;
        try
        {
            using var doc = JsonDocument.Parse(body);
            var root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object) return null;
            if (root.Str("message") is { Length: > 0 } msg) return msg.Trim();
            if (root.Prop("errors") is { ValueKind: JsonValueKind.Object } errs)
                return string.Join("; ", errs.EnumerateObject().Select(p => $"{p.Name}: {p.Value}"));
            return null;
        }
        catch (JsonException)
        {
            return body.Length > 200 ? body[..200] : body;
        }
    }

    public void Dispose()
    {
        _http.Dispose();
        _gate.Dispose();
    }

    /// <summary>A handler for real connections: per-collection, with optional certificate bypass.</summary>
    public static HttpMessageHandler CreateDefaultHandler(bool ignoreCertificateErrors)
    {
        var handler = new SocketsHttpHandler
        {
            ConnectTimeout = TimeSpan.FromSeconds(10),
            UseCookies = false,
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
            PooledConnectionLifetime = TimeSpan.FromMinutes(5),
            MaxConnectionsPerServer = 16,
        };
        if (ignoreCertificateErrors)
            handler.SslOptions.RemoteCertificateValidationCallback = (_, _, _, _) => true;
        return handler;
    }
}
