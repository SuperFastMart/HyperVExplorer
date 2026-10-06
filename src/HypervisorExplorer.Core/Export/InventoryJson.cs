using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Core.Export;

/// <summary>
/// Saves and loads inventory snapshots as JSON, so an estate can be captured once and reopened,
/// compared, or exported later without reconnecting.
/// </summary>
public static partial class InventoryJson
{
    public const string Format = "HypervisorExplorer.Inventory";
    public const int Version = 1;

    public static readonly JsonSerializerOptions Options = CreateOptions();

    private static JsonSerializerOptions CreateOptions()
    {
        var o = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        };
        o.Converters.Add(new JsonStringEnumConverter());
        o.Converters.Add(new LooseObjectConverter());
        return o;
    }

    private sealed class Document
    {
        public string Format { get; set; } = InventoryJson.Format;
        public int Version { get; set; } = InventoryJson.Version;
        public DateTimeOffset ExportedAt { get; set; } = DateTimeOffset.Now;
        public string? Application { get; set; } = "Hypervisor Explorer";
        public List<InventorySnapshot> Snapshots { get; set; } = [];
    }

    public static void Save(IEnumerable<InventorySnapshot> snapshots, string path)
    {
        var doc = new Document { Snapshots = snapshots.ToList() };
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        JsonSerializer.Serialize(fs, doc, Options);
    }

    public static string Serialize(IEnumerable<InventorySnapshot> snapshots) =>
        JsonSerializer.Serialize(new Document { Snapshots = snapshots.ToList() }, Options);

    public static List<InventorySnapshot> Load(string path)
    {
        using var fs = File.OpenRead(path);
        return Deserialize(fs);
    }

    public static List<InventorySnapshot> Deserialize(Stream stream)
    {
        var doc = JsonSerializer.Deserialize<Document>(stream, Options)
                  ?? throw new InvalidDataException("Empty inventory file.");
        if (!string.Equals(doc.Format, Format, StringComparison.Ordinal))
            throw new InvalidDataException($"Not a Hypervisor Explorer inventory file (format '{doc.Format}').");
        if (doc.Version > Version)
            throw new InvalidDataException($"Inventory file version {doc.Version} is newer than this app supports ({Version}).");
        return doc.Snapshots;
    }

    public static List<InventorySnapshot> Deserialize(string json)
    {
        using var ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(json));
        return Deserialize(ms);
    }

    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}")]
    private static partial Regex IsoDate();

    /// <summary>
    /// Round-trips <c>object?</c> values (the Extra bags) as plain CLR types instead of JsonElement:
    /// numbers become long/double, ISO timestamps DateTime, arrays of objects List&lt;Dictionary&gt;.
    /// </summary>
    private sealed class LooseObjectConverter : JsonConverter<object>
    {
        public override object? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            using var doc = JsonDocument.ParseValue(ref reader);
            return Convert(doc.RootElement);
        }

        private static object? Convert(JsonElement e) => e.ValueKind switch
        {
            JsonValueKind.Null or JsonValueKind.Undefined => null,
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Number => e.TryGetInt64(out var l) ? l : e.GetDouble(),
            JsonValueKind.String => e.GetString() is { } s && IsoDate().IsMatch(s) && DateTime.TryParse(s, null,
                System.Globalization.DateTimeStyles.RoundtripKind, out var dt) ? dt : e.GetString(),
            JsonValueKind.Object => e.EnumerateObject()
                .ToDictionary(p => p.Name, p => Convert(p.Value), StringComparer.OrdinalIgnoreCase),
            JsonValueKind.Array => ConvertArray(e),
            _ => e.GetRawText(),
        };

        private static object ConvertArray(JsonElement e)
        {
            var items = e.EnumerateArray().Select(Convert).ToList();
            return items.Count > 0 && items.All(i => i is Dictionary<string, object?>)
                ? items.Cast<Dictionary<string, object?>>().ToList()
                : items;
        }

        public override void Write(Utf8JsonWriter writer, object value, JsonSerializerOptions options)
        {
            if (value.GetType() == typeof(object))
            {
                writer.WriteStartObject();
                writer.WriteEndObject();
                return;
            }
            JsonSerializer.Serialize(writer, value, value.GetType(), options);
        }
    }
}
