using System.Dynamic;

namespace wedgeautodraw_1_2.Core.Models;

public class NamedDimensionAnnotations
{
    private readonly Dictionary<string, DimensionAnnotation> _data = new();

    public DimensionAnnotation this[string key]
    {
        get => _data.ContainsKey(key) ? _data[key] : null;
        set => _data[key] = value;
    }

    public Dictionary<string, DimensionAnnotation> GetAll() => _data;
    public bool ContainsKey(string key) => _data.ContainsKey(key);
    public bool TryGet(string key, out DimensionAnnotation value)
    {
        return _data.TryGetValue(key, out value);
    }


}
