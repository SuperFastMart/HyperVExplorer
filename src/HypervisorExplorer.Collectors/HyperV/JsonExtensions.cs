using System.Globalization;
using System.Text.Json;

namespace HypervisorExplorer.Collectors.HyperV;

/// <summary>
/// Lenient accessors over the collection JSON. Windows PowerShell 5.1 can unroll one-element arrays or emit
/// numbers as strings, so every accessor tolerates both shapes.
/// </summary>
internal static class JsonExtensions
{
    public static JsonElement? Prop(this JsonElement e, string name) =>
        e.ValueKind == JsonValueKind.Object && e.TryGetProperty(name, out var v) && v.ValueKind != JsonValueKind.Null ? v : null;

    public static JsonElement? Prop(this JsonElement? e, string name) => e is { } x ? x.Prop(name) : null;

    public static string? Str(this JsonElement e, string name) => e.Prop(name) is { } v ? AsString(v) : null;

    public static string? Str(this JsonElement? e, string name) => e is { } x ? x.Str(name) : null;

    public static long? Long(this JsonElement e, string name) => e.Prop(name) is { } v ? AsLong(v) : null;

    public static long? Long(this JsonElement? e, string name) => e is { } x ? x.Long(name) : null;

    public static int? Int(this JsonElement e, string name) => e.Long(name) is { } l && l is >= int.MinValue and <= int.MaxValue ? (int)l : null;

    public static int? Int(this JsonElement? e, string name) => e is { } x ? x.Int(name) : null;

    public static bool? Bool(this JsonElement e, string name)
    {
        if (e.Prop(name) is not { } v) return null;
        return v.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Number => v.TryGetInt64(out var n) ? n != 0 : null,
            JsonValueKind.String => bool.TryParse(v.GetString(), out var b) ? b : null,
            _ => null,
        };
    }

    public static bool? Bool(this JsonElement? e, string name) => e is { } x ? x.Bool(name) : null;

    public static DateTimeOffset? Date(this JsonElement e, string name)
    {
        var s = e.Str(name);
        if (s is null) return null;
        if (DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var d) && d.Year > 1900)
            return d;
        return null;
    }

    public static DateTimeOffset? Date(this JsonElement? e, string name) => e is { } x ? x.Date(name) : null;

    /// <summary>Array items; a lone object/value is treated as a one-element array.</summary>
    public static IEnumerable<JsonElement> Arr(this JsonElement e, string name)
    {
        if (e.Prop(name) is not { } v) yield break;
        if (v.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in v.EnumerateArray())
                if (item.ValueKind != JsonValueKind.Null) yield return item;
        }
        else
        {
            yield return v;
        }
    }

    public static IEnumerable<JsonElement> Arr(this JsonElement? e, string name) => e is { } x ? x.Arr(name) : [];

    public static List<string> Strs(this JsonElement e, string name) =>
        e.Arr(name).Select(AsString).Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s!).ToList();

    public static List<string> Strs(this JsonElement? e, string name) => e is { } x ? x.Strs(name) : [];

    private static string? AsString(JsonElement v) => v.ValueKind switch
    {
        JsonValueKind.String => v.GetString() is { Length: > 0 } s ? s : null,
        JsonValueKind.Number => v.GetRawText(),
        JsonValueKind.True => "True",
        JsonValueKind.False => "False",
        _ => null,
    };

    private static long? AsLong(JsonElement v)
    {
        switch (v.ValueKind)
        {
            case JsonValueKind.Number:
                if (v.TryGetInt64(out var l)) return l;
                if (v.TryGetDouble(out var d) && d is >= long.MinValue and <= long.MaxValue) return (long)Math.Round(d);
                return null;
            case JsonValueKind.String:
                var s = v.GetString();
                if (long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var p)) return p;
                if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var pd) && pd is >= long.MinValue and <= long.MaxValue)
                    return (long)Math.Round(pd);
                return null;
            default:
                return null;
        }
    }
}
