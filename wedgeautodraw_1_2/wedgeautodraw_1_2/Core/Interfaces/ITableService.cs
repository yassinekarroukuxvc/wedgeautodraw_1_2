using wedgeautodraw_1_2.Core.Models;
namespace wedgeautodraw_1_2.Core.Interfaces;

public interface ITableService
{
    bool CreateDimensionTable(DataStorage position, string[] wedgeKeys, string header, DrawingData drawingData, NamedDimensionValues wedgeDimensions);
    bool CreateLabelAsTable(DataStorage position, DrawingData drawingData);
    bool CreatePolishTable(DataStorage position, DrawingData drawingData);
    bool CreateHowToOrderTable(DataStorage position, string header, DrawingData drawingData);
}
