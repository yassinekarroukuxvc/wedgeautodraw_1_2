using System.Dynamic;

namespace wedgeautodraw_1_2.Core.Models;

public class DynamicDataContainer : DynamicObject
{
    private readonly Dictionary<string, DataStorage> _data = new();

    public DataStorage this[string key]
    {
        get => _data.ContainsKey(key) ? _data[key] : null;
        set => _data[key] = value;
    }

    public Dictionary<string, DataStorage> GetAll() => _data;

    public bool ContainsKey(string key) => _data.ContainsKey(key);

    public bool TryGet(string key, out DataStorage value)
    {
        return _data.TryGetValue(key, out value);
    }

}
