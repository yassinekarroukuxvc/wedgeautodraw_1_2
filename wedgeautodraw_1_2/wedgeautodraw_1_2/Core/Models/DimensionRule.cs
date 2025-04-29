namespace wedgeautodraw_1_2.Core.Models;

public class DimensionRule
{
    public string BasedOnView { get; set; }
    public Func<WedgeData, DrawingData, double[]> CalculatePosition { get; set; }
}