using System.Text.Json;
using wedgeautodraw_1_2.Core.Enums;

namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public class ConfigLoader
{
    private readonly Dictionary<string, JsonElement> _config;

    public ConfigLoader(string jsonPath, DrawingType drawingType)
    {
        if (!File.Exists(jsonPath))
        {
            throw new FileNotFoundException("Configuration JSON file not found.", jsonPath);
        }

        var json = File.ReadAllText(jsonPath);
        var fullConfig = JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, JsonElement>>>(json);

        string typeKey = drawingType.ToString();
        if (!fullConfig.TryGetValue(typeKey, out var section))
        {
            throw new InvalidOperationException($"Drawing type section '{typeKey}' not found in config.");
        }

        _config = section;
    }

    public double GetDouble(string key)
    {
        if (_config.TryGetValue(key, out var element))
        {
            if (element.ValueKind == JsonValueKind.Number && element.TryGetDouble(out double val))
                return val;

            if (element.ValueKind == JsonValueKind.String && double.TryParse(element.GetString(), out val))
                return val;
        }

        return double.NaN;
    }

    public string GetString(string key)
    {
        if (_config.TryGetValue(key, out var element))
        {
            return element.ToString();
        }

        return string.Empty;
    }

    public string[] GetStringArray(string key, char separator = '¶')
    {
        var raw = GetString(key);
        return raw.Split(separator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
    }

    public bool GetBool(string key)
    {
        if (_config.TryGetValue(key, out var element))
        {
            if (element.ValueKind == JsonValueKind.True) return true;
            if (element.ValueKind == JsonValueKind.False) return false;

            if (element.ValueKind == JsonValueKind.String && bool.TryParse(element.GetString(), out bool result))
                return result;
        }

        return false;
    }
    public bool HasKey(string key)
    {
        return _config.ContainsKey(key);
    }
}
