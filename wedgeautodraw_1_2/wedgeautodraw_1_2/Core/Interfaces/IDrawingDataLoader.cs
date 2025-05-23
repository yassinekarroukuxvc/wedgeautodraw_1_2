using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Core.Interfaces;

public interface IDrawingDataLoader
{
    DrawingData LoadDrawingData(WedgeData wedgeData, string configFilePath, DrawingType drawingType);

}
