using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services;

public static class DynamicDimensionStyler
{
    public static void ApplyDynamicStyles(DrawingData drawingData, WedgeData wedgeData)
    {
        foreach (var kvp in DimensionRules.Rules)
        {
            string dimName = kvp.Key;
            var rule = kvp.Value;

            double[] computedPosition = rule.CalculatePosition(wedgeData, drawingData);

            if (computedPosition != null && computedPosition.Length == 2)
            {
                drawingData.DimensionStyles[dimName] = new DimensionAnnotation(new DataStorage(computedPosition));
            }
        }
    }
}
