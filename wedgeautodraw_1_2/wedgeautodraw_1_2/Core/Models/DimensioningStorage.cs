using SolidWorks.Interop.swconst;

namespace wedgeautodraw_1_2.Core.Models;

public class DimensioningStorage
{
    public DataStorage Position { get; set; }
    public DimensioningStorage(DataStorage position)
    {
        Position = position;
    }
}
