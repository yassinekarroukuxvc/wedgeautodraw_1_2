using SolidWorks.Interop.swconst;

namespace wedgeautodraw_1_2.Core.Models;

public class DimensionAnnotation
{
    public DataStorage Position { get; set; }
    public DimensionAnnotation(DataStorage position)
    {
        Position = position;
    }
}
