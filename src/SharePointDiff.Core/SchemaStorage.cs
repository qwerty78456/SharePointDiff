using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Collections.Concurrent;

namespace SharePointDiff.Core;

/// <summary>Immutable, content-addressed schemas. Legacy inline arrays remain readable.</summary>
internal sealed class SchemaStorage : IDisposable
{
    private static readonly AsyncLocal<SchemaStorage?> Active = new();
    private readonly SchemaStorage? previous;
    private readonly string? root;
    private readonly Dictionary<string, JsonElement> schemas = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<JsonElement, string> ids;
    private readonly ConcurrentDictionary<string, byte> verified;
    private readonly ConcurrentDictionary<string, JsonElement> catalog;
    private SchemaStorage(string? root)
    {
        previous = Active.Value; this.root = root;
        ids = previous != null && previous.root == root ? previous.ids : new();
        verified = previous != null && previous.root == root ? previous.verified : new();
        catalog = previous != null && previous.root == root ? previous.catalog : new();
        Active.Value = this;
    }
    internal static SchemaStorage Begin(Workspace workspace) => new(Path.Combine(workspace.DataRoot, "schemas"));
    public static SchemaStorage For(string path)
    {
        var directory = new DirectoryInfo(Path.GetDirectoryName(Path.GetFullPath(path))!);
        while (directory != null && directory.Parent != null)
        {
            string candidate = Path.Combine(directory.FullName, "schemas");
            if (Directory.Exists(candidate)) return new(candidate);
            directory = directory.Parent;
        }
        return new(null);
    }
    internal static JsonElement ReadSchema(string path, string id)
    {
        using var storage = For(path);
        return storage.Load(id);
    }
    internal static string Hash(JsonElement schema)
    {
        using var bytes = new MemoryStream();
        using (var writer = new Utf8JsonWriter(bytes)) Canonical(writer, schema);
        return Convert.ToHexString(SHA256.HashData(bytes.ToArray())).ToLowerInvariant();
    }
    private static void Canonical(Utf8JsonWriter writer, JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Object)
        {
            writer.WriteStartObject();
            foreach (var property in value.EnumerateObject().OrderBy(p => p.Name, StringComparer.Ordinal))
            { writer.WritePropertyName(property.Name); Canonical(writer, property.Value); }
            writer.WriteEndObject();
        }
        else if (value.ValueKind == JsonValueKind.Array)
        {
            writer.WriteStartArray();
            foreach (var entry in value.EnumerateArray()) Canonical(writer, entry);
            writer.WriteEndArray();
        }
        else value.WriteTo(writer);
    }
    public async Task PersistAsync(CancellationToken ct)
    {
        if (root == null) return;
        foreach (var (id, schema) in schemas)
        {
            string path = Path.Combine(root, id + ".json");
            if (File.Exists(path))
            {
                if (!verified.ContainsKey(id))
                {
                    if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) throw new OperationException("MetadataStoreCorrupt", path);
                    ReadValidated(path, id);
                    verified.TryAdd(id, 0);
                }
                catalog.TryAdd(id, schema);
                continue;
            }
            try { await JsonFiles.WriteNewAsync(path, schema, ct); }
            catch (IOException) when (File.Exists(path))
            {
                // Two transfer workers can publish the same immutable schema concurrently.
                var existing = JsonFiles.Read<JsonElement>(path);
                if (existing.ValueKind != JsonValueKind.Array || Hash(existing) != id) throw new OperationException("MetadataStoreCorrupt", path);
            }
            verified.TryAdd(id, 0);
            catalog.TryAdd(id, schema);
        }
    }
    private JsonElement Load(string id)
    {
        if (root == null || id.Length != 64 || id.Any(c => !char.IsAsciiHexDigit(c))) throw new OperationException("MetadataStoreCorrupt", root);
        if (catalog.TryGetValue(id, out var cached)) return cached;
        string path = Path.Combine(root, id + ".json");
        if (!File.Exists(path) || (File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) throw new OperationException("MetadataStoreCorrupt", path);
        var schema = ReadValidated(path, id);
        verified.TryAdd(id, 0);
        catalog.TryAdd(id, schema);
        return schema;
    }
    private static JsonElement ReadValidated(string path, string id)
    {
        try
        {
            IoMetrics.Add("Metadata.Reads"); IoMetrics.Add("Metadata.ReadBytes", new FileInfo(path).Length);
            using var json = JsonDocument.Parse(File.ReadAllText(path));
            if (json.RootElement.ValueKind != JsonValueKind.Array || Hash(json.RootElement) != id) throw new OperationException("MetadataStoreCorrupt", path);
            return json.RootElement.Clone();
        }
        catch (JsonException ex) { throw new OperationException("MetadataStoreCorrupt", path, ex); }
    }
    public void Dispose() => Active.Value = previous;
    internal sealed class MetadataConverter : JsonConverter<MetadataSnapshot>
    {
        public override MetadataSnapshot Read(ref Utf8JsonReader reader, Type type, JsonSerializerOptions options)
        {
            using var doc = JsonDocument.ParseValue(ref reader); var item = doc.RootElement;
            var columns = item.GetProperty("Columns");
            bool shared = columns.ValueKind == JsonValueKind.Object && columns.TryGetProperty("SchemaRef", out _);
            if (shared)
                columns = Active.Value?.Load(columns.GetProperty("SchemaRef").GetString()!) ?? throw new OperationException("MetadataStoreCorrupt");
            return new(item.GetProperty("Fields").Clone(), item.GetProperty("DisplayValues").Clone(), shared ? columns : columns.Clone(),
                item.GetProperty("Complete").GetBoolean(), item.TryGetProperty("Error", out var error) && error.ValueKind == JsonValueKind.String ? error.GetString() : null,
                item.TryGetProperty("ObservedAt", out var observed) && observed.ValueKind == JsonValueKind.String ? observed.GetDateTimeOffset() : null);
        }
        public override void Write(Utf8JsonWriter writer, MetadataSnapshot value, JsonSerializerOptions options)
        {
            writer.WriteStartObject(); writer.WritePropertyName("Fields"); value.Fields.WriteTo(writer);
            writer.WritePropertyName("DisplayValues"); value.DisplayValues.WriteTo(writer);
            writer.WritePropertyName("Columns");
            if (Active.Value is { root: not null } storage)
            {
                if (!storage.ids.TryGetValue(value.Columns, out string? id)) storage.ids[value.Columns] = id = Hash(value.Columns);
                storage.schemas.TryAdd(id, value.Columns);
                writer.WriteStartObject(); writer.WriteString("SchemaRef", id); writer.WriteEndObject();
            }
            else value.Columns.WriteTo(writer);
            writer.WriteBoolean("Complete", value.Complete); writer.WriteString("Error", value.Error);
            if (value.ObservedAt is { } observed) writer.WriteString("ObservedAt", observed);
            writer.WriteEndObject();
        }
    }
    internal sealed class ManifestDictionaryConverter : JsonConverter<Dictionary<string, InstalledRecord>>
    {
        public override Dictionary<string, InstalledRecord> Read(ref Utf8JsonReader reader, Type type, JsonSerializerOptions options)
        {
            using var doc = JsonDocument.ParseValue(ref reader); var root = doc.RootElement;
            if (root.TryGetProperty("FormatVersion", out _) && root.TryGetProperty("Records", out var records)) root = records;
            var result = new Dictionary<string, InstalledRecord>(StringComparer.Ordinal);
            foreach (var pair in root.EnumerateObject()) result.Add(pair.Name, pair.Value.Deserialize<InstalledRecord>(options)!);
            return result;
        }
        public override void Write(Utf8JsonWriter writer, Dictionary<string, InstalledRecord> value, JsonSerializerOptions options)
        {
            writer.WriteStartObject();
            foreach (var (key, record) in value) { writer.WritePropertyName(key); JsonSerializer.Serialize(writer, record, options); }
            writer.WriteEndObject();
        }
    }
}
