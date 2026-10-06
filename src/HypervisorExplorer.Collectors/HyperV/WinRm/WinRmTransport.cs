using System.Buffers;
using System.Buffers.Binary;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Security;
using System.Text;

namespace HypervisorExplorer.Collectors.HyperV.WinRm;

public enum WinRmErrorKind
{
    Unreachable,
    AuthenticationFailed,
    AccessDenied,
    Certificate,
    Protocol,
}

/// <summary>A WinRM transport or WS-Management failure.</summary>
public sealed class WinRmException(WinRmErrorKind kind, string message, Exception? inner = null) : Exception(message, inner)
{
    public WinRmErrorKind Kind { get; } = kind;
}

/// <summary>
/// Sends WS-Management SOAP messages to a Windows host, authenticating with NTLM (username/password).
/// Over HTTP (5985) every message is sealed with the NTLM session key using WinRM's
/// "application/HTTP-SPNEGO-session-encrypted" framing, so the default WinRM configuration
/// (AllowUnencrypted = false) is accepted without changes on the host. Over HTTPS (5986) TLS protects
/// the channel and bodies are sent as plain SOAP.
/// </summary>
/// <remarks>
/// NTLM authenticates a TCP connection, so the client is pinned to a single pooled connection and requests
/// are serialised. If the server drops the connection, the next request is re-authenticated automatically.
/// </remarks>
public sealed class WinRmTransport : IDisposable
{
    private const string Boundary = "Encrypted Boundary";
    private const string EncryptedProtocol = "application/HTTP-SPNEGO-session-encrypted";
    private static readonly byte[] OctetStreamMarker = Encoding.ASCII.GetBytes("\tContent-Type: application/octet-stream\r\n");
    private static readonly byte[] EndBoundary = Encoding.ASCII.GetBytes("--" + Boundary + "--\r\n");

    private readonly HttpClient _http;
    private readonly NetworkCredential _credential;
    private readonly string _host;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private NegotiateAuthentication? _auth;
    private string _scheme = "Negotiate";

    static WinRmTransport()
    {
        // Use .NET's managed NTLM on macOS/Linux so behaviour doesn't depend on the OS GSSAPI build.
        AppContext.SetSwitch("System.Net.Security.UseManagedNtlm", true);
    }

    public WinRmTransport(string host, int port, bool useHttps, string username, string password,
        bool ignoreCertificateErrors, TimeSpan requestTimeout, HttpMessageHandler? handler = null)
    {
        _host = host;
        UseHttps = useHttps;
        var uriHost = host.Contains(':') && !host.StartsWith('[') ? $"[{host}]" : host;
        Endpoint = new Uri($"{(useHttps ? "https" : "http")}://{uriHost}:{port}/wsman");
        _credential = ParseCredential(username, password);

        if (handler is null)
        {
            var sockets = new SocketsHttpHandler
            {
                MaxConnectionsPerServer = 1,
                PooledConnectionLifetime = Timeout.InfiniteTimeSpan,
                PooledConnectionIdleTimeout = TimeSpan.FromMinutes(5),
                ConnectTimeout = TimeSpan.FromSeconds(15),
                UseCookies = false,
                UseProxy = false,
                AllowAutoRedirect = false,
            };
            if (ignoreCertificateErrors)
                sockets.SslOptions.RemoteCertificateValidationCallback = (_, _, _, _) => true;
            handler = sockets;
        }
        _http = new HttpClient(handler, disposeHandler: true) { Timeout = requestTimeout };
    }

    public Uri Endpoint { get; }
    public bool UseHttps { get; }
    private bool Encrypt => !UseHttps;

    /// <summary>"DOMAIN\user" → domain + user; "user@domain" (UPN) and plain local names are passed as-is.</summary>
    internal static NetworkCredential ParseCredential(string username, string password)
    {
        var slash = username.IndexOf('\\');
        return slash > 0
            ? new NetworkCredential(username[(slash + 1)..], password, username[..slash])
            : new NetworkCredential(username, password, "");
    }

    /// <summary>Posts a SOAP envelope and returns the SOAP response (decrypted). Throws <see cref="WinRmException"/>.</summary>
    public async Task<string> SendAsync(string soap, CancellationToken ct)
    {
        await _gate.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            for (var attempt = 0; ; attempt++)
            {
                if (_auth is null) await AuthenticateAsync(ct).ConfigureAwait(false);
                var (status, text) = await PostAsync(Encoding.UTF8.GetBytes(soap), ct).ConfigureAwait(false);
                if (status == HttpStatusCode.Unauthorized && attempt == 0)
                {
                    // The authenticated connection was dropped (idle timeout, server restart): log in again once.
                    ResetAuth();
                    continue;
                }
                return status switch
                {
                    HttpStatusCode.OK => text,
                    HttpStatusCode.Unauthorized => throw new WinRmException(WinRmErrorKind.AuthenticationFailed,
                        "The WinRM session was rejected after re-authenticating."),
                    HttpStatusCode.InternalServerError when text.Contains("Fault", StringComparison.Ordinal) =>
                        throw WsManFault.Parse(text),
                    _ => throw new WinRmException(WinRmErrorKind.Protocol,
                        $"WinRM returned HTTP {(int)status} {status}.{(text.Length > 0 ? " " + Truncate(text) : "")}"),
                };
            }
        }
        finally
        {
            _gate.Release();
        }
    }

    private void ResetAuth()
    {
        _auth?.Dispose();
        _auth = null;
    }

    private async Task AuthenticateAsync(CancellationToken ct)
    {
        var auth = new NegotiateAuthentication(new NegotiateAuthenticationClientOptions
        {
            Package = "NTLM",
            Credential = _credential,
            TargetName = "HTTP/" + _host,
            RequiredProtectionLevel = Encrypt ? ProtectionLevel.EncryptAndSign : ProtectionLevel.None,
        });
        try
        {
            var negotiate = auth.GetOutgoingBlob(ReadOnlySpan<byte>.Empty, out var status);
            if (negotiate is null || status != NegotiateAuthenticationStatusCode.ContinueNeeded)
                throw new WinRmException(WinRmErrorKind.AuthenticationFailed, $"Could not start NTLM authentication ({status}).");

            using var first = await SendAuthAsync("Negotiate", negotiate, ct).ConfigureAwait(false);
            if (first.StatusCode != HttpStatusCode.Unauthorized)
                throw new WinRmException(WinRmErrorKind.Protocol,
                    $"Expected an authentication challenge from {Endpoint} but got HTTP {(int)first.StatusCode}.");

            var offered = first.Headers.WwwAuthenticate.ToList();
            var challenge = offered.FirstOrDefault(h =>
                (h.Scheme.Equals("Negotiate", StringComparison.OrdinalIgnoreCase) || h.Scheme.Equals("NTLM", StringComparison.OrdinalIgnoreCase))
                && !string.IsNullOrEmpty(h.Parameter));
            if (challenge is null)
            {
                var schemes = string.Join(", ", offered.Select(h => h.Scheme).Distinct());
                throw new WinRmException(WinRmErrorKind.AuthenticationFailed,
                    $"The host did not accept NTLM/Negotiate authentication (offered: {(schemes.Length > 0 ? schemes : "none")}).");
            }
            _scheme = challenge.Scheme;

            var authenticate = auth.GetOutgoingBlob(Convert.FromBase64String(challenge.Parameter!), out status);
            if (authenticate is null || status is not (NegotiateAuthenticationStatusCode.Completed or NegotiateAuthenticationStatusCode.ContinueNeeded))
                throw new WinRmException(WinRmErrorKind.AuthenticationFailed, $"NTLM authentication failed ({status}).");

            using var second = await SendAuthAsync(_scheme, authenticate, ct).ConfigureAwait(false);
            if (second.StatusCode == HttpStatusCode.Unauthorized)
                throw new WinRmException(WinRmErrorKind.AuthenticationFailed, "The user name or password was rejected.");
            _auth = auth;
        }
        catch
        {
            auth.Dispose();
            throw;
        }
    }

    /// <summary>An empty POST carrying an NTLM token (the same handshake pywinrm uses before sealed messages).</summary>
    private async Task<HttpResponseMessage> SendAuthAsync(string scheme, byte[] token, CancellationToken ct)
    {
        var request = new HttpRequestMessage(HttpMethod.Post, Endpoint)
        {
            Content = new ByteArrayContent([]),
        };
        request.Headers.Authorization = new AuthenticationHeaderValue(scheme, Convert.ToBase64String(token));
        request.Content.Headers.ContentType = Encrypt
            ? MediaTypeHeaderValue.Parse($"multipart/encrypted;protocol=\"{EncryptedProtocol}\";boundary=\"{Boundary}\"")
            : MediaTypeHeaderValue.Parse("application/soap+xml;charset=UTF-8");
        var response = await SendRawAsync(request, ct).ConfigureAwait(false);
        await response.Content.ReadAsByteArrayAsync(ct).ConfigureAwait(false); // drain so the connection is reused
        return response;
    }

    private async Task<(HttpStatusCode Status, string Text)> PostAsync(byte[] body, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, Endpoint);
        if (Encrypt)
        {
            request.Content = new ByteArrayContent(Seal(body));
            request.Content.Headers.ContentType =
                MediaTypeHeaderValue.Parse($"multipart/encrypted;protocol=\"{EncryptedProtocol}\";boundary=\"{Boundary}\"");
        }
        else
        {
            request.Content = new ByteArrayContent(body);
            request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/soap+xml;charset=UTF-8");
        }

        using var response = await SendRawAsync(request, ct).ConfigureAwait(false);
        var bytes = await response.Content.ReadAsByteArrayAsync(ct).ConfigureAwait(false);
        if (bytes.Length == 0) return (response.StatusCode, "");
        var mediaType = response.Content.Headers.ContentType?.MediaType ?? "";
        var plain = mediaType.Equals("multipart/encrypted", StringComparison.OrdinalIgnoreCase) ? Unseal(bytes) : bytes;
        return (response.StatusCode, Encoding.UTF8.GetString(plain));
    }

    private async Task<HttpResponseMessage> SendRawAsync(HttpRequestMessage request, CancellationToken ct)
    {
        try
        {
            return await _http.SendAsync(request, HttpCompletionOption.ResponseContentRead, ct).ConfigureAwait(false);
        }
        catch (HttpRequestException ex) when (ex.InnerException is System.Security.Authentication.AuthenticationException)
        {
            throw new WinRmException(WinRmErrorKind.Certificate, $"TLS handshake with {Endpoint} failed: {ex.InnerException.Message}", ex);
        }
        catch (HttpRequestException ex)
        {
            throw new WinRmException(WinRmErrorKind.Unreachable, $"Cannot reach {Endpoint}: {ex.Message}", ex);
        }
        catch (TaskCanceledException ex) when (!ct.IsCancellationRequested)
        {
            throw new WinRmException(WinRmErrorKind.Unreachable, $"{Endpoint} did not respond within {_http.Timeout.TotalSeconds:0} s.", ex);
        }
    }

    /// <summary>Seals a SOAP body into WinRM's multipart/encrypted framing.</summary>
    internal byte[] Seal(byte[] body)
    {
        var writer = new ArrayBufferWriter<byte>(body.Length + 64);
        var status = _auth!.Wrap(body, writer, requestEncryption: true, out var encrypted);
        if (status != NegotiateAuthenticationStatusCode.Completed || !encrypted)
            throw new WinRmException(WinRmErrorKind.Protocol, $"Could not encrypt the WinRM message ({status}).");
        return Frame(writer.WrittenSpan, body.Length);
    }

    /// <summary>NTLM wrap output is a 16-byte signature followed by the sealed data.</summary>
    internal static byte[] Frame(ReadOnlySpan<byte> wrapped, int originalLength, int signatureLength = 16)
    {
        var ms = new MemoryStream(wrapped.Length + 256);
        void Ascii(string s) => ms.Write(Encoding.ASCII.GetBytes(s));
        Ascii($"--{Boundary}\r\n");
        Ascii($"\tContent-Type: {EncryptedProtocol}\r\n");
        Ascii($"\tOriginalContent: type=application/soap+xml;charset=UTF-8;Length={originalLength}\r\n");
        Ascii($"--{Boundary}\r\n");
        ms.Write(OctetStreamMarker);
        Span<byte> len = stackalloc byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(len, signatureLength);
        ms.Write(len);
        ms.Write(wrapped);
        ms.Write(EndBoundary);
        return ms.ToArray();
    }

    private byte[] Unseal(byte[] framed)
    {
        var wrapped = Unframe(framed, out var expectedLength);
        var writer = new ArrayBufferWriter<byte>(wrapped.Length);
        var status = _auth!.Unwrap(wrapped, writer, out _);
        if (status != NegotiateAuthenticationStatusCode.Completed)
            throw new WinRmException(WinRmErrorKind.Protocol, $"Could not decrypt the WinRM response ({status}).");
        var plain = writer.WrittenSpan.ToArray();
        if (expectedLength is { } n && n != plain.Length)
            throw new WinRmException(WinRmErrorKind.Protocol, $"Decrypted WinRM response has {plain.Length} bytes; expected {n}.");
        return plain;
    }

    /// <summary>Extracts signature+sealed data from the multipart framing (signature length prefix removed).</summary>
    internal static byte[] Unframe(byte[] framed, out int? expectedLength)
    {
        expectedLength = null;
        var headerText = Encoding.ASCII.GetString(framed, 0, Math.Min(framed.Length, 512));
        var lengthAt = headerText.IndexOf("Length=", StringComparison.Ordinal);
        if (lengthAt >= 0)
        {
            var digits = new string(headerText[(lengthAt + 7)..].TakeWhile(char.IsDigit).ToArray());
            if (int.TryParse(digits, out var n)) expectedLength = n;
        }

        var start = framed.AsSpan().IndexOf(OctetStreamMarker);
        if (start < 0) throw new WinRmException(WinRmErrorKind.Protocol, "Malformed encrypted WinRM response (no payload part).");
        var payload = framed.AsSpan(start + OctetStreamMarker.Length);
        var end = payload.LastIndexOf(EndBoundary);
        if (end >= 0) payload = payload[..end];
        if (payload.Length < 4) throw new WinRmException(WinRmErrorKind.Protocol, "Malformed encrypted WinRM response (truncated).");
        var sigLen = BinaryPrimitives.ReadInt32LittleEndian(payload);
        if (sigLen <= 0 || sigLen > payload.Length - 4)
            throw new WinRmException(WinRmErrorKind.Protocol, "Malformed encrypted WinRM response (bad signature length).");
        return payload[4..].ToArray();
    }

    private static string Truncate(string s) => s.Length <= 300 ? s : s[..300] + "…";

    public void Dispose()
    {
        _auth?.Dispose();
        _http.Dispose();
        _gate.Dispose();
    }
}
