using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Core.Enums;
using System;
using System.Collections.Generic;

namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public static class DynamicDimensionRuleUpdater
{
    public static void OverrideRuleIf(this Dictionary<string, DimensionLayoutRule> rules,
                                      string dimensionKey,
                                      Func<WedgeData, DrawingData, bool> condition,
                                      Func<WedgeData, DrawingData, double[]> newPosition)
    {
        if (!rules.ContainsKey(dimensionKey))
            return;

        var originalRule = rules[dimensionKey];
        rules[dimensionKey] = new DimensionLayoutRule
        {
            BasedOnView = originalRule.BasedOnView,
            CalculatePosition = (wedge, drawing) =>
            {
                if (condition(wedge, drawing))
                    return newPosition(wedge, drawing);
                return originalRule.CalculatePosition(wedge, drawing);
            }
        };
    }

    public static void InjectRule(this Dictionary<string, DimensionLayoutRule> rules,
                                  string dimensionKey,
                                  string viewName,
                                  Func<WedgeData, DrawingData, double[]> positionLogic)
    {
        rules[dimensionKey] = new DimensionLayoutRule
        {
            BasedOnView = viewName,
            CalculatePosition = positionLogic
        };
    }
}
