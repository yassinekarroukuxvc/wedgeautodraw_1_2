using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Core.Interfaces;

public interface IDataContainerLoader
{

    WedgeData LoadWedgeData(string toleranceFilePath);
    DrawingData LoadDrawingData(WedgeData wedgeData, string configFilePath);
}
