using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Core.Enums;
using System;
using System.Collections.Generic;

namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public static class DrawingDataUpdater
{
    public static void UpdateIf(this NamedDimensionValues values, string key, Func<DataStorage, bool> condition, Func<DataStorage> newValueFactory)
    {
        if (values.TryGet(key, out var existingValue) && condition(existingValue))
        {
            values[key] = newValueFactory();
        }
    }

    public static void UpdateValue(this NamedDimensionValues values, string key, Func<DataStorage, DataStorage> transformer)
    {
        if (values.TryGet(key, out var existingValue))
        {
            values[key] = transformer(existingValue);
        }
    }

    public static void AddOrUpdate(this NamedDimensionValues values, string key, DataStorage value)
    {
        values[key] = value;
    }

    public static void UpdateTitleBlockInfo(this DrawingData data, string key, Func<string, string> updater)
    {
        if (data.TitleBlockInfo.TryGetValue(key, out var val))
        {
            data.TitleBlockInfo[key] = updater(val);
        }
    }

    public static void SetIfMissing(this NamedDimensionValues values, string key, DataStorage defaultValue)
    {
        if (!values.ContainsKey(key))
        {
            values[key] = defaultValue;
        }
    }
}
