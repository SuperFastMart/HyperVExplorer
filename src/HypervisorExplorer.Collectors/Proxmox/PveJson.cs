using System.Globalization;
using System.Text.Json;

namespace HypervisorExplorer.Collectors.Proxmox;

/// <summary>
/// Tolerant accessors for PVE API JSON. The API is loosely typed: the same field can arrive as a number in one
/// endpoint/version and a string in another ("mtu": "9000", "mhz": "2400.000", "onboot": 1).
/// </summary>
internal static class PveJson
{
    public static bool Has(this JsonElement e, string name) =>
        e.ValueKind == JsonValueKind.Object && e.TryGetProperty(name, out var v) && v.ValueKind != JsonValueKind.Null;

    public static JsonElement? Prop(this JsonElement e, string name) =>
        e.ValueKind == JsonValueKind.Object && e.TryGetProperty(name, out var v) && v.ValueKind != JsonValueKind.Null
            ? v
            : null;

    public static string? Str(this JsonElement e, string name)
    {
        if (e.Prop(name) is not { } v) return null;
        return v.ValueKind switch
        {
            JsonValueKind.String => v.GetString(),
            JsonValueKind.Number => v.GetRawText(),
            JsonValueKind.True => "1",
            JsonValueKind.False => "0",
            _ => v.GetRawText(),
        };
    }

    public static double? Dbl(this JsonElement e, string name)
    {
        if (e.Prop(name) is not { } v) return null;
        if (v.ValueKind == JsonValueKind.Number) return v.GetDouble();
        if (v.ValueKind == JsonValueKind.String &&
            double.TryParse(v.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return d;
        if (v.ValueKind == JsonValueKind.True) return 1;
        if (v.ValueKind == JsonValueKind.False) return 0;
        return null;
    }

    public static long? Long(this JsonElement e, string name)
    {
        if (e.Prop(name) is not { } v) return null;
        if (v.ValueKind == JsonValueKind.Number)
            return v.TryGetInt64(out var l) ? l : (long)Math.Round(v.GetDouble());
        return Dbl(e, name) is { } d ? (long)Math.Round(d) : null;
    }

    public static int? Int(this JsonElement e, string name) => Long(e, name) is { } l ? (int)l : null;

    /// <summary>PVE booleans are 0/1 integers, occasionally strings or JSON booleans.</summary>
    public static bool? Bool(this JsonElement e, string name)
    {
        if (e.Prop(name) is not { } v) return null;
        return v.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Number => v.GetDouble() != 0,
            JsonValueKind.String => PveParsing.IsTrue(v.GetString()),
            _ => null,
        };
    }

    public static IEnumerable<JsonElement> Items(this JsonElement? e) =>
        e is { ValueKind: JsonValueKind.Array } a ? a.EnumerateArray() : [];

    public static IEnumerable<JsonElement> Items(this JsonElement e) =>
        e.ValueKind == JsonValueKind.Array ? e.EnumerateArray() : [];

    /// <summary>Enumerates object properties as (name, string value) for config objects.</summary>
    public static IEnumerable<(string Key, string Value)> StringProps(this JsonElement e)
    {
        if (e.ValueKind != JsonValueKind.Object) yield break;
        foreach (var p in e.EnumerateObject())
        {
            if (p.Value.ValueKind == JsonValueKind.Null) continue;
            yield return (p.Name, e.Str(p.Name) ?? "");
        }
    }
}
