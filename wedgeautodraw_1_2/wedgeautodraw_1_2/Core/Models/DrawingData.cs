using wedgeautodraw_1_2.Core.Enums;

namespace wedgeautodraw_1_2.Core.Models;

public class DrawingData
{
    public NamedDimensionValues ViewPositions { get; set; } = new();
    public NamedDimensionValues ViewScales { get; set; } = new();
    public NamedDimensionValues TablePositions { get; set; } = new();
    public NamedDimensionValues BreaklineData { get; set; } = new();

    public Dictionary<string, string> TitleBlockInfo { get; set; } = new();
    public Dictionary<string, string> HowToOrderInfo { get; set; } = new();
    public Dictionary<string, string> TitleInfo { get; set; } = new();

    public string[] LabelAsItems { get; set; } = new string[] { };
    public string[] PolishItems { get; set; } = new string[] { };
    public string[] DimensionKeysInTable { get; set; } = new string[] { };

    public NamedDimensionAnnotations DimensionStyles { get; set; } = new();
    public DrawingType DrawingType { get; set; }
    public string Title { get; set; } = string.Empty;
    public override string ToString()
    {
        var result = new System.Text.StringBuilder();
        result.AppendLine("=== DrawingData ===");

        result.AppendLine("\nTitle: " + Title);
        result.AppendLine("DrawingType: " + DrawingType);

        result.AppendLine("\nTitleInfo:");
        foreach (var kvp in TitleInfo)
            result.AppendLine($"  {kvp.Key}: {kvp.Value}");

        result.AppendLine("\nTitleBlockInfo:");
        foreach (var kvp in TitleBlockInfo)
            result.AppendLine($"  {kvp.Key}: {kvp.Value}");

        result.AppendLine("\nHowToOrderInfo:");
        foreach (var kvp in HowToOrderInfo)
            result.AppendLine($"  {kvp.Key}: {kvp.Value}");

        result.AppendLine("\nLabelAsItems:");
        foreach (var item in LabelAsItems)
            result.AppendLine($"  - {item}");

        result.AppendLine("\nPolishItems:");
        foreach (var item in PolishItems)
            result.AppendLine($"  - {item}");

        result.AppendLine("\nDimensionKeysInTable:");
        foreach (var key in DimensionKeysInTable)
            result.AppendLine($"  - {key}");

        result.AppendLine("\nViewPositions:");
        foreach (var kvp in ViewPositions.GetAll())
            result.AppendLine($"  {kvp.Key}: {string.Join(", ", kvp.Value.GetValues(Unit.Millimeter))} mm");

        result.AppendLine("\nViewScales:");
        foreach (var kvp in ViewScales.GetAll())
            result.AppendLine($"  {kvp.Key}: {kvp.Value.GetValue(Unit.Millimeter)}");

        result.AppendLine("\nTablePositions:");
        foreach (var kvp in TablePositions.GetAll())
            result.AppendLine($"  {kvp.Key}: {string.Join(", ", kvp.Value.GetValues(Unit.Millimeter))}");

        result.AppendLine("\nBreaklineData:");
        foreach (var kvp in BreaklineData.GetAll())
            result.AppendLine($"  {kvp.Key}: {kvp.Value.GetValue(Unit.Millimeter)} mm");

        result.AppendLine("\nDimensionStyles:");
        foreach (var kvp in DimensionStyles.GetAll())
        {
            var style = kvp.Value;
            result.AppendLine($"  {kvp.Key}: Pos = ({string.Join(", ", style.Position?.GetValues(Unit.Millimeter) ?? new double[0])})");
        }

        return result.ToString();
    }

}
