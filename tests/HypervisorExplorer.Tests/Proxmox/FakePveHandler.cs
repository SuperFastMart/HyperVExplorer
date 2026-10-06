using System.Collections.Concurrent;
using System.Net;
using System.Text;

namespace HypervisorExplorer.Tests.Proxmox;

/// <summary>A recorded request (headers captured before the message is disposed).</summary>
public sealed record RecordedRequest(string Method, string Path, string? Authorization, string? Cookie, string? Csrf, string? Body);

/// <summary>Serves fixture JSON by API path (relative to /api2/json, including any query string).</summary>
public sealed class FakePveHandler : HttpMessageHandler
{
    private readonly ConcurrentDictionary<string, Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>>> _routes = new(StringComparer.Ordinal);

    public ConcurrentQueue<RecordedRequest> Requests { get; } = new();

    /// <summary>Thrown for every request when set (simulates transport failures).</summary>
    public Exception? ThrowOnSend { get; set; }

    public FakePveHandler Data(string path, string dataJson)
    {
        _routes[path] = (_, _) => Task.FromResult(Json($"{{\"data\":{dataJson}}}"));
        return this;
    }

    public FakePveHandler Status(string path, HttpStatusCode code, string reason)
    {
        _routes[path] = (_, _) => Task.FromResult(new HttpResponseMessage(code)
        {
            ReasonPhrase = reason,
            Content = new StringContent("{\"data\":null}", Encoding.UTF8, "application/json"),
        });
        return this;
    }

    public FakePveHandler Route(string path, Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> responder)
    {
        _routes[path] = responder;
        return this;
    }

    public static HttpResponseMessage Json(string body) =>
        new(HttpStatusCode.OK) { Content = new StringContent(body, Encoding.UTF8, "application/json") };

    public IEnumerable<string> Paths => Requests.Select(r => r.Path);

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var path = request.RequestUri!.PathAndQuery;
        const string prefix = "/api2/json";
        if (path.StartsWith(prefix, StringComparison.Ordinal)) path = path[prefix.Length..];
        string? Header(string name) => request.Headers.TryGetValues(name, out var v) ? string.Join(";", v) : null;
        var body = request.Content is null ? null : await request.Content.ReadAsStringAsync(cancellationToken);
        Requests.Enqueue(new RecordedRequest(request.Method.Method, path, Header("Authorization"), Header("Cookie"),
            Header("CSRFPreventionToken"), body));

        if (ThrowOnSend is not null) throw ThrowOnSend;
        if (_routes.TryGetValue(path, out var responder)) return await responder(request, cancellationToken);
        return new HttpResponseMessage(HttpStatusCode.NotImplemented)
        {
            ReasonPhrase = $"Method '{request.Method} {path}' not implemented",
            Content = new StringContent("{\"data\":null}"),
        };
    }
}
