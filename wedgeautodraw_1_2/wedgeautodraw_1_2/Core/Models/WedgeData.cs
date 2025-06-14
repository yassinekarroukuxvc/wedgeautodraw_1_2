using wedgeautodraw_1_2.Core.Enums;

namespace wedgeautodraw_1_2.Core.Models;

public class WedgeData
{
    public NamedDimensionValues Dimensions { get; set; } = new();
    public Dictionary<string, object> Metadata { get; set; } = new();
    public string EngravedText { get; set; } = string.Empty;
    public string OverlayCalibration { get; set; } = string.Empty;
    public double OverlayScaling { get; set; } = 1.0;
    public override string ToString()
    {
        var result = new System.Text.StringBuilder();
        result.AppendLine("=== WedgeData ===");

        result.AppendLine("Dimensions:");
        foreach (var kvp in Dimensions.GetAll())
        {
            var value = kvp.Value.GetValue(Unit.Millimeter);
            var tolPlus = kvp.Value.GetTolerance(Unit.Millimeter, "+");
            var tolMinus = kvp.Value.GetTolerance(Unit.Millimeter, "-");
            result.AppendLine($"  {kvp.Key}: {value:F3} mm (+{tolPlus:F3}/-{tolMinus:F3})");
        }

        result.AppendLine("\nMetadata:");
        foreach (var meta in Metadata)
        {
            result.AppendLine($"  {meta.Key}: {meta.Value}");
        }

        result.AppendLine($"\nEngravedText: {EngravedText}");
        result.AppendLine($"\nOverlay Calibration: {OverlayCalibration}");
        result.AppendLine($"\nOverlay Scaling: {OverlayScaling}");
        return result.ToString();
    }

}
