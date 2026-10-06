using System.Globalization;
using System.Security;
using System.Xml.Linq;

namespace HypervisorExplorer.Collectors.VMware;

/// <summary>Helpers for reading vim25 (urn:vim25) SOAP payloads with LINQ to XML.</summary>
internal static class VimXml
{
    public static readonly XNamespace Vim = "urn:vim25";
    public static readonly XNamespace Xsi = "http://www.w3.org/2001/XMLSchema-instance";
    public static readonly XNamespace Soap = "http://schemas.xmlsoap.org/soap/envelope/";

    /// <summary>First vim25 child element with the given local name.</summary>
    public static XElement? El(this XElement? e, string name) => e?.Element(Vim + name);

    /// <summary>All vim25 child elements with the given local name (array members of a data object).</summary>
    public static IEnumerable<XElement> Els(this XElement? e, string name) =>
        e?.Elements(Vim + name) ?? Enumerable.Empty<XElement>();

    /// <summary>Navigates a dotted path of single-valued children, e.g. "shares.level".</summary>
    public static XElement? At(this XElement? e, string path)
    {
        foreach (var part in path.Split('.'))
        {
            if (e is null) return null;
            e = e.El(part);
        }
        return e;
    }

    public static string? Str(this XElement? e, string? path = null)
    {
        var t = path is null ? e : e.At(path);
        return t?.Value;
    }

    public static int? Int(this XElement? e, string? path = null) =>
        int.TryParse(e.Str(path), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null;

    public static long? Long(this XElement? e, string? path = null) =>
        long.TryParse(e.Str(path), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null;

    public static double? Dbl(this XElement? e, string? path = null) =>
        double.TryParse(e.Str(path), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : null;

    public static bool? Bool(this XElement? e, string? path = null) => e.Str(path) switch
    {
        "true" or "1" => true,
        "false" or "0" => false,
        _ => null,
    };

    public static DateTimeOffset? Date(this XElement? e, string? path = null) =>
        DateTimeOffset.TryParse(e.Str(path), CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var v)
            ? v : null;

    /// <summary>Local part of the xsi:type attribute (e.g. "VirtualVmxnet3"), or null.</summary>
    public static string? XsiType(this XElement? e)
    {
        var t = e?.Attribute(Xsi + "type")?.Value;
        if (t is null) return null;
        var i = t.IndexOf(':');
        return i >= 0 ? t[(i + 1)..] : t;
    }

    public static List<string> Strs(this XElement? e, string name) => e.Els(name).Select(x => x.Value).ToList();

    public static string Esc(string? s) => SecurityElement.Escape(s ?? "") ?? "";
}

/// <summary>A managed object reference (type + value, e.g. VirtualMachine:vm-42).</summary>
public sealed record MoRef(string Type, string Value)
{
    internal static MoRef? From(XElement? e) =>
        e is null || string.IsNullOrEmpty(e.Value) ? null : new MoRef(e.Attribute("type")?.Value ?? "", e.Value.Trim());

    internal string ToXml(string tag) => $"<{tag} type=\"{VimXml.Esc(Type)}\">{VimXml.Esc(Value)}</{tag}>";

    public override string ToString() => $"{Type}:{Value}";
}

/// <summary>One ObjectContent from RetrievePropertiesEx: a managed object and its requested property values.</summary>
public sealed class VimObject
{
    public VimObject(MoRef reference) => Ref = reference;

    public MoRef Ref { get; }

    /// <summary>Property path (as requested) → the &lt;val&gt; element.</summary>
    public Dictionary<string, XElement> Props { get; } = new(StringComparer.Ordinal);

    /// <summary>Property paths reported in missingSet (e.g. no permission).</summary>
    public List<string> Missing { get; } = [];

    /// <summary>
    /// Returns the value at <paramref name="path"/>; when only a parent path was retrieved (e.g. "guest"),
    /// navigates into it ("guest.toolsStatus").
    /// </summary>
    public XElement? Get(string path)
    {
        if (Props.TryGetValue(path, out var v)) return v;
        var best = "";
        foreach (var key in Props.Keys)
        {
            if (key.Length > best.Length && path.Length > key.Length && path.StartsWith(key, StringComparison.Ordinal)
                && path[key.Length] == '.')
                best = key;
        }
        return best.Length == 0 ? null : Props[best].At(path[(best.Length + 1)..]);
    }

    public string? Str(string path) => Get(path)?.Value;
    public bool? Bool(string path) => Get(path).Bool();
    public int? Int(string path) => Get(path).Int();
    public long? Long(string path) => Get(path).Long();
    public DateTimeOffset? Date(string path) => Get(path).Date();
    public MoRef? Ref1(string path) => MoRef.From(Get(path));

    /// <summary>Array-valued property (ArrayOfX wrapper) or repeated children of a nested path.</summary>
    public IEnumerable<XElement> Items(string path)
    {
        if (Props.TryGetValue(path, out var v)) return v.Elements();
        var i = path.LastIndexOf('.');
        if (i < 0) return [];
        return Get(path[..i]).Els(path[(i + 1)..]);
    }

    public List<MoRef> Refs(string path) => Items(path).Select(MoRef.From).OfType<MoRef>().ToList();
}
