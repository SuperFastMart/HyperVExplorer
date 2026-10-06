using System.Net;
using System.Net.Sockets;
using System.Security.Authentication;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using HypervisorExplorer.Core.Collection;
using static HypervisorExplorer.Collectors.VMware.VimXml;

namespace HypervisorExplorer.Collectors.VMware;

/// <summary>The parts of the vSphere ServiceContent the collector uses.</summary>
public sealed class ServiceContent
{
    public required XElement About { get; init; }
    public required MoRef RootFolder { get; init; }
    public required MoRef PropertyCollector { get; init; }
    public required MoRef ViewManager { get; init; }
    public required MoRef SessionManager { get; init; }
    public MoRef? LicenseManager { get; init; }

    public string Name => About.Str("name") ?? "VMware";
    public string FullName => About.Str("fullName") ?? Name;
    public string ApiType => About.Str("apiType") ?? "";
    public string ApiVersion => About.Str("apiVersion") ?? "";
    public string Version => About.Str("version") ?? "";
    public string? Build => About.Str("build");
    public string? InstanceUuid => About.Str("instanceUuid");
    public bool IsVCenter => ApiType == "VirtualCenter";
}

/// <summary>A SOAP fault returned by the vSphere API (vim25 MethodFault subtype).</summary>
public sealed class VimFaultException : Exception
{
    public VimFaultException(string faultType, string message, XElement? detail)
        : base(message)
    {
        FaultType = faultType;
        Detail = detail;
    }

    /// <summary>The vim25 fault type, e.g. "InvalidLogin", "NoPermission", "InvalidProperty".</summary>
    public string FaultType { get; }

    /// <summary>The fault detail element (e.g. &lt;InvalidLoginFault xsi:type="InvalidLogin"/&gt;).</summary>
    public XElement? Detail { get; }
}

/// <summary>
/// Minimal vSphere Web Services (SOAP, vim25) client over <see cref="HttpClient"/>: the same API RVTools uses,
/// against <c>https://host/sdk</c> on ESXi and vCenter. Keeps the <c>vmware_soap_session</c> cookie.
/// </summary>
public sealed class VSphereClient : IDisposable
{
    private readonly HttpClient _http;
    private readonly CookieContainer _cookies = new();
    private string _soapAction = "urn:vim25";

    public VSphereClient(Uri sdkUri, HttpMessageHandler handler, bool disposeHandler = true)
    {
        SdkUri = sdkUri;
        _http = new HttpClient(handler, disposeHandler) { Timeout = TimeSpan.FromMinutes(10) };
    }

    public Uri SdkUri { get; }

    public ServiceContent? Content { get; private set; }

    /// <summary>Creates a client for a connection request; <paramref name="handler"/> overrides the transport (tests).</summary>
    public static VSphereClient Create(ConnectionRequest request, HttpMessageHandler? handler = null)
    {
        var host = request.Address.Trim();
        if (host.Contains(':') && !host.StartsWith('[') && IPAddress.TryParse(host, out var ip)
            && ip.AddressFamily == AddressFamily.InterNetworkV6)
            host = $"[{host}]";
        var uri = new UriBuilder("https", host, request.Port ?? 443, "/sdk").Uri;
        if (handler is not null) return new VSphereClient(uri, handler);

        var sockets = new SocketsHttpHandler
        {
            UseCookies = false,
            ConnectTimeout = TimeSpan.FromSeconds(20),
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
            PooledConnectionLifetime = TimeSpan.FromMinutes(5),
        };
        if (request.IgnoreCertificateErrors)
            sockets.SslOptions.RemoteCertificateValidationCallback = (_, _, _, _) => true;
        return new VSphereClient(uri, sockets);
    }

    public void Dispose() => _http.Dispose();

    // ---------------------------------------------------------------- operations

    public async Task<ServiceContent> RetrieveServiceContentAsync(CancellationToken ct)
    {
        var resp = await InvokeAsync("RetrieveServiceContent", "<_this type=\"ServiceInstance\">ServiceInstance</_this>", ct);
        var rv = resp.El("returnval") ?? throw Protocol("RetrieveServiceContent returned no ServiceContent.");
        var about = rv.El("about") ?? throw Protocol("ServiceContent has no 'about' information.");
        Content = new ServiceContent
        {
            About = about,
            RootFolder = MoRef.From(rv.El("rootFolder")) ?? throw Protocol("ServiceContent has no rootFolder."),
            PropertyCollector = MoRef.From(rv.El("propertyCollector")) ?? throw Protocol("ServiceContent has no propertyCollector."),
            ViewManager = MoRef.From(rv.El("viewManager")) ?? throw Protocol("ServiceContent has no viewManager (vSphere 4.0 or later is required)."),
            SessionManager = MoRef.From(rv.El("sessionManager")) ?? throw Protocol("ServiceContent has no sessionManager."),
            LicenseManager = MoRef.From(rv.El("licenseManager")),
        };
        if (!string.IsNullOrEmpty(Content.ApiVersion)) _soapAction = "urn:vim25/" + Content.ApiVersion;
        return Content;
    }

    public async Task LoginAsync(string userName, string password, CancellationToken ct)
    {
        var sc = RequireContent();
        await InvokeAsync("Login",
            sc.SessionManager.ToXml("_this") + $"<userName>{Esc(userName)}</userName><password>{Esc(password)}</password>", ct);
    }

    public Task LogoutAsync(CancellationToken ct) =>
        InvokeAsync("Logout", RequireContent().SessionManager.ToXml("_this"), ct);

    public async Task<MoRef> CreateContainerViewAsync(MoRef container, IEnumerable<string> types, bool recursive, CancellationToken ct)
    {
        var sc = RequireContent();
        var body = new StringBuilder(sc.ViewManager.ToXml("_this")).Append(container.ToXml("container"));
        foreach (var t in types) body.Append("<type>").Append(Esc(t)).Append("</type>");
        body.Append("<recursive>").Append(recursive ? "true" : "false").Append("</recursive>");
        var resp = await InvokeAsync("CreateContainerView", body.ToString(), ct);
        return MoRef.From(resp.El("returnval")) ?? throw Protocol("CreateContainerView returned no view.");
    }

    public Task DestroyViewAsync(MoRef view, CancellationToken ct) => InvokeAsync("DestroyView", view.ToXml("_this"), ct);

    /// <summary>
    /// Retrieves <paramref name="paths"/> for every object of <paramref name="type"/> in a ContainerView, paging with
    /// ContinueRetrievePropertiesEx. Paths the server rejects (InvalidProperty, e.g. newer than its API version)
    /// are dropped with a warning and the call is retried.
    /// </summary>
    public async Task<List<VimObject>> RetrieveFromViewAsync(MoRef view, string type, IReadOnlyCollection<string> paths,
        int maxObjects, Action<int>? onPage, ICollection<string>? warnings, CancellationToken ct)
    {
        var objectSet = view.ToXml("obj") + "<skip>true</skip><selectSet xsi:type=\"TraversalSpec\"><name>traverseView</name>"
            + "<type>ContainerView</type><path>view</path><skip>false</skip></selectSet>";
        return await RetrieveAsync(type, paths, objectSet, maxObjects, onPage, warnings, ct);
    }

    /// <summary>Retrieves properties of a single managed object (no traversal).</summary>
    public async Task<VimObject?> RetrieveObjectAsync(MoRef obj, IReadOnlyCollection<string> paths, ICollection<string>? warnings,
        CancellationToken ct)
    {
        var list = await RetrieveAsync(obj.Type, paths, obj.ToXml("obj") + "<skip>false</skip>", 0, null, warnings, ct);
        return list.FirstOrDefault();
    }

    private async Task<List<VimObject>> RetrieveAsync(string type, IReadOnlyCollection<string> paths, string objectSetXml,
        int maxObjects, Action<int>? onPage, ICollection<string>? warnings, CancellationToken ct)
    {
        var sc = RequireContent();
        var remaining = paths.Distinct().ToList();
        for (var attempt = 0; ; attempt++)
        {
            var body = new StringBuilder(sc.PropertyCollector.ToXml("_this"));
            body.Append("<specSet><propSet><type>").Append(Esc(type)).Append("</type><all>false</all>");
            foreach (var p in remaining) body.Append("<pathSet>").Append(Esc(p)).Append("</pathSet>");
            body.Append("</propSet><objectSet>").Append(objectSetXml).Append("</objectSet></specSet><options>");
            if (maxObjects > 0) body.Append("<maxObjects>").Append(maxObjects).Append("</maxObjects>");
            body.Append("</options>");

            XElement resp;
            try
            {
                resp = await InvokeAsync("RetrievePropertiesEx", body.ToString(), ct);
            }
            catch (VimFaultException f) when (f.FaultType == "InvalidProperty" && attempt < 20)
            {
                var bad = f.Detail.Str("name");
                var drop = bad is null ? [] : remaining.Where(p => PathMatches(p, bad)).ToList();
                if (drop.Count == 0) throw;
                foreach (var d in drop) remaining.Remove(d);
                warnings?.Add($"{type}: property '{string.Join("', '", drop)}' is not supported by this server and was skipped.");
                if (remaining.Count == 0) return [];
                continue;
            }

            var results = new List<VimObject>();
            var token = ParseResult(resp.El("returnval"), results);
            onPage?.Invoke(results.Count);
            while (!string.IsNullOrEmpty(token))
            {
                ct.ThrowIfCancellationRequested();
                var next = await InvokeAsync("ContinueRetrievePropertiesEx",
                    sc.PropertyCollector.ToXml("_this") + $"<token>{Esc(token)}</token>", ct);
                token = ParseResult(next.El("returnval"), results);
                onPage?.Invoke(results.Count);
            }
            return results;
        }
    }

    private static bool PathMatches(string path, string bad) =>
        path == bad || path.StartsWith(bad + ".", StringComparison.Ordinal)
        || path.EndsWith("." + bad, StringComparison.Ordinal) || path.Contains("." + bad + ".", StringComparison.Ordinal);

    /// <summary>Parses a RetrieveResult into <paramref name="into"/>; returns the continuation token, if any.</summary>
    internal static string? ParseResult(XElement? result, List<VimObject> into)
    {
        if (result is null) return null;
        foreach (var oc in result.Els("objects"))
        {
            var r = MoRef.From(oc.El("obj"));
            if (r is null) continue;
            var o = new VimObject(r);
            foreach (var ps in oc.Els("propSet"))
            {
                var name = ps.Str("name");
                var val = ps.El("val");
                if (name is not null && val is not null) o.Props[name] = val;
            }
            foreach (var ms in oc.Els("missingSet"))
            {
                if (ms.Str("path") is { } p) o.Missing.Add(p);
            }
            into.Add(o);
        }
        return result.Str("token");
    }

    /// <summary>Invokes a vim25 method and returns its &lt;{op}Response&gt; element. Throws <see cref="VimFaultException"/>
    /// for SOAP faults and <see cref="CollectionException"/> for transport failures.</summary>
    public async Task<XElement> InvokeAsync(string operation, string innerXml, CancellationToken ct)
    {
        var envelope = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" "
            + "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
            + $"<soapenv:Body><{operation} xmlns=\"urn:vim25\">{innerXml}</{operation}></soapenv:Body></soapenv:Envelope>";

        using var req = new HttpRequestMessage(HttpMethod.Post, SdkUri)
        {
            Content = new StringContent(envelope, new UTF8Encoding(false), "text/xml"),
        };
        req.Content.Headers.ContentType!.CharSet = "utf-8";
        req.Headers.TryAddWithoutValidation("SOAPAction", _soapAction);
        var cookie = _cookies.GetCookieHeader(SdkUri);
        if (!string.IsNullOrEmpty(cookie)) req.Headers.TryAddWithoutValidation("Cookie", cookie);

        HttpResponseMessage resp;
        try
        {
            resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or IOException)
        {
            throw TransportError(ex);
        }

        using (resp)
        {
            if (resp.Headers.TryGetValues("Set-Cookie", out var setCookies))
            {
                foreach (var sc in setCookies)
                {
                    try { _cookies.SetCookies(SdkUri, sc); }
                    catch (CookieException) { /* ignore malformed cookies other than the session */ }
                }
            }

            XDocument doc;
            try
            {
                await using var stream = await resp.Content.ReadAsStreamAsync(ct);
                doc = await XDocument.LoadAsync(stream, LoadOptions.None, ct);
            }
            catch (XmlException ex)
            {
                throw new CollectionException(CollectionFailure.Protocol,
                    $"{SdkUri} returned HTTP {(int)resp.StatusCode} {resp.ReasonPhrase} with a non-SOAP response.",
                    "Check that the address is an ESXi host or vCenter Server and that the vSphere Web Services SDK (/sdk) is reachable.",
                    ex);
            }
            catch (Exception ex) when (ex is HttpRequestException or IOException)
            {
                throw TransportError(ex);
            }

            var body = doc.Root?.Element(Soap + "Body")
                ?? throw Protocol($"{SdkUri} returned HTTP {(int)resp.StatusCode} without a SOAP body.");
            if (body.Element(Soap + "Fault") is { } fault) throw ParseFault(fault);
            if (!resp.IsSuccessStatusCode)
                throw Protocol($"{operation} failed: HTTP {(int)resp.StatusCode} {resp.ReasonPhrase}.");
            return body.Element(Vim + operation + "Response") ?? body.Elements().FirstOrDefault()
                ?? throw Protocol($"{operation} returned an empty SOAP body.");
        }
    }

    internal static VimFaultException ParseFault(XElement fault)
    {
        var message = fault.Element("faultstring")?.Value ?? fault.Element(Soap + "faultstring")?.Value ?? "SOAP fault";
        var detail = (fault.Element("detail") ?? fault.Element(Soap + "detail"))?.Elements().FirstOrDefault();
        var type = detail.XsiType() ?? detail?.Name.LocalName;
        if (type is not null && type.EndsWith("Fault", StringComparison.Ordinal) && detail.XsiType() is null)
            type = type[..^"Fault".Length];
        return new VimFaultException(type ?? "Unknown", message.Trim(), detail);
    }

    private CollectionException TransportError(Exception ex)
    {
        var host = SdkUri.Host;
        var port = SdkUri.Port;
        for (var e = ex; e is not null; e = e.InnerException)
        {
            if (e is AuthenticationException)
                return new CollectionException(CollectionFailure.Certificate,
                    $"TLS handshake with {host}:{port} failed: {e.Message}",
                    "The server certificate is not trusted. Enable 'Ignore certificate errors' (RVTools does this by default) or install a trusted certificate on the ESXi host / vCenter.",
                    ex);
            if (e is SocketException se)
                return new CollectionException(CollectionFailure.Unreachable,
                    $"Cannot connect to {host}:{port}: {se.Message}",
                    $"Check the name/IP and that TCP port {port} (HTTPS, vSphere Web Services /sdk) is open from this machine.", ex);
        }
        if (ex is TaskCanceledException)
            return new CollectionException(CollectionFailure.Unreachable, $"Timed out talking to {host}:{port}.",
                $"Check that {host} is reachable on TCP port {port} (HTTPS) and responsive.", ex);
        return new CollectionException(CollectionFailure.Unreachable, $"Cannot reach {SdkUri}: {ex.Message}",
            $"Check the name/IP and that TCP port {port} (HTTPS) is open from this machine.", ex);
    }

    /// <summary>Maps a vSphere fault to a user-facing <see cref="CollectionException"/>.</summary>
    public static CollectionException ToCollectionException(VimFaultException f, string? userName = null) => f.FaultType switch
    {
        "InvalidLogin" or "InvalidLocale" => new CollectionException(CollectionFailure.AuthenticationFailed,
            $"Login failed{(userName is null ? "" : " for '" + userName + "'")}: {f.Message}",
            "Use a vCenter SSO account like administrator@vsphere.local (or a domain account in UPN form, user@domain) "
            + "or an ESXi local account (root). A read-only role is sufficient.", f),
        "NotAuthenticated" => new CollectionException(CollectionFailure.AuthenticationFailed,
            $"The vSphere session is not authenticated: {f.Message}",
            "The session expired or was terminated. Reconnect; check that the account is not locked.", f),
        "NoPermission" => new CollectionException(CollectionFailure.PermissionDenied,
            $"Permission denied: {f.Message}" + (f.Detail.Str("privilegeId") is { } p ? $" (privilege {p})" : ""),
            "Grant the account at least the Read-only role at the vCenter root (propagate to children) or on the ESXi host.", f),
        "NoPermissionOnHost" or "RestrictedVersion" => new CollectionException(CollectionFailure.PermissionDenied,
            f.Message,
            "Free ESXi (vSphere Hypervisor) licences block API access; a licensed host or vCenter is required.", f),
        _ => new CollectionException(CollectionFailure.Protocol, $"vSphere API fault {f.FaultType}: {f.Message}", null, f),
    };

    private ServiceContent RequireContent() =>
        Content ?? throw new InvalidOperationException("Call RetrieveServiceContentAsync first.");

    private static CollectionException Protocol(string message) =>
        new(CollectionFailure.Protocol, message, "The server did not respond like a vSphere SDK endpoint.");
}
