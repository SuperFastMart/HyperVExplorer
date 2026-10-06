using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using static HypervisorExplorer.Tests.VMware.VSphereFixtures;

namespace HypervisorExplorer.Tests.VMware;

/// <summary>
/// Fake vSphere /sdk endpoint: routes SOAP requests by body operation (and, for RetrievePropertiesEx, by the
/// PropertySpec type) to fixture responses. Records every call for assertions.
/// </summary>
internal sealed partial class FakeVSphereHandler : HttpMessageHandler
{
    public const string SessionCookie = "vmware_soap_session=\"52b5f1d2-0d36-6a40-5d84-8b4f1f0c9e1a\"";

    public sealed record Call(string Operation, string Body, string? SoapAction, string? Cookie, Uri? Uri);

    public List<Call> Calls { get; } = [];

    /// <summary>Optional override: return a response body (and status) for an operation/type, or null to use defaults.</summary>
    public Func<string, string?, string, (HttpStatusCode, string)?>? Override { get; set; }

    public int Count(string op) => Calls.Count(c => c.Operation == op);

    [GeneratedRegex(@"<soapenv:Body><(\w+)")]
    private static partial Regex OpRegex();

    [GeneratedRegex(@"<propSet><type>(\w+)</type>")]
    private static partial Regex PropTypeRegex();

    [GeneratedRegex(@"<type>(\w+)</type>")]
    private static partial Regex TypeRegex();

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var body = await request.Content!.ReadAsStringAsync(cancellationToken);
        var op = OpRegex().Match(body).Groups[1].Value;
        request.Headers.TryGetValues("SOAPAction", out var actions);
        request.Headers.TryGetValues("Cookie", out var cookies);
        Calls.Add(new Call(op, body, actions?.FirstOrDefault(), cookies?.FirstOrDefault(), request.RequestUri));

        var type = op switch
        {
            "RetrievePropertiesEx" => PropTypeRegex().Match(body).Groups[1].Value,
            "CreateContainerView" => TypeRegex().Match(body).Groups[1].Value,
            _ => null,
        };

        var (status, xml) = Override?.Invoke(op, type, body) ?? Default(op, type, body);
        var resp = new HttpResponseMessage(status) { Content = new StringContent(xml, Encoding.UTF8, "text/xml") };
        if (op == "Login" && status == HttpStatusCode.OK)
            resp.Headers.TryAddWithoutValidation("Set-Cookie", SessionCookie + "; Path=/; HttpOnly; Secure;");
        return resp;
    }

    private static (HttpStatusCode, string) Ok(string xml) => (HttpStatusCode.OK, xml);

    public static (HttpStatusCode, string) ServerFault(string type, string message, string detail = "") =>
        (HttpStatusCode.InternalServerError, Fault(type, message, detail));

    private static (HttpStatusCode, string) Default(string op, string? type, string body) => op switch
    {
        "RetrieveServiceContent" => Ok(Envelope(ServiceContent)),
        "Login" => body.Contains("<password>wrong</password>")
            ? ServerFault("InvalidLogin", "Cannot complete login due to an incorrect user name or password.")
            : Ok(Envelope(LoginResponse)),
        "Logout" => Ok(Empty("Logout")),
        "CreateContainerView" => Ok(ContainerView(type ?? "")),
        "DestroyView" => Ok(Empty("DestroyView")),
        "RetrievePropertiesEx" => type switch
        {
            "ManagedEntity" => Ok(Entities()),
            "HostSystem" => Ok(Hosts()),
            "ComputeResource" => Ok(ComputeResources()),
            "ResourcePool" => Ok(ResourcePools()),
            "Datastore" => Ok(Datastores()),
            "DistributedVirtualSwitch" => Ok(Dvs()),
            "DistributedVirtualPortgroup" => Ok(Portgroups()),
            // First page of VMs: two objects and a continuation token.
            "VirtualMachine" => Ok(RetrieveResult(Web01 + Db01, token: "1")),
            "LicenseManager" => Ok(Licenses()),
            _ => Ok(RetrieveResult("")),
        },
        "ContinueRetrievePropertiesEx" => Ok(RetrieveResult(Template, continuation: true)),
        "QueryAssignedLicenses" => Ok(AssignedLicenses()),
        _ => ServerFault("MethodNotFound", $"Method {op} not found"),
    };
}
