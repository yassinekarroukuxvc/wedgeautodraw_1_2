using System.Dynamic;

namespace wedgeautodraw_1_2.Core.Models;

public class DynamicDimensioningContainer : DynamicObject
{
    private readonly Dictionary<string, DimensioningStorage> _data = new();

    public DimensioningStorage this[string key]
    {
        get => _data.ContainsKey(key) ? _data[key] : null;
        set => _data[key] = value;
    }

    public Dictionary<string, DimensioningStorage> GetAll() => _data;
}
