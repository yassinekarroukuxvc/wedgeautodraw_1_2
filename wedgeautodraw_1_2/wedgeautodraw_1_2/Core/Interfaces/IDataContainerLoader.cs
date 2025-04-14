using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Core.Interfaces;

public interface IDataContainerLoader
{
   
    WedgeData LoadWedgeData();
    DrawingData LoadDrawingData(WedgeData wedgeData, string configFilePath);
}
