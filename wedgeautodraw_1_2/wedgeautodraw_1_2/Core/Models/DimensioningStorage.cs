using SolidWorks.Interop.swconst;

namespace wedgeautodraw_1_2.Core.Models;

public class DimensioningStorage
{
    public DataStorage Position { get; set; }
    public int? ArrowSide { get; set; }
    public bool? CenterText { get; set; }
    public int? WitnessVisibility { get; set; }
    public bool? ExtensionLineFromCenter { get; set; }
    public ExtensionLineAsCenterline ExtensionLineAsCenterline { get; set; }

    public DimensioningStorage(DataStorage position)
    {
        Position = position;
    }
}

public class ExtensionLineAsCenterline
{
    public short? ExtIndex { get; set; }
    public bool? Centerline { get; set; }

}
