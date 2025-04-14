using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Infrastructure.Services;

public class TableService : ITableService
{
    private readonly ModelDoc2 _swModel;
    private readonly DrawingDoc _swDrawing;
    private readonly SldWorks _swApp;

    public TableService(SldWorks swApp, ModelDoc2 swModel)
    {
        _swApp = swApp;
        _swModel = swModel;
        _swDrawing = (DrawingDoc)swModel;
    }

    public bool CreateDimensionTable(DataStorage position, string[] wedgeKeys, string header, DrawingData drawingData, DynamicDataContainer wedgeDimensions)
    {
        try
        {
            var table = (TableAnnotation)_swDrawing.InsertTableAnnotation2(false, position.GetValues(Unit.Meter)[0], position.GetValues(Unit.Meter)[1], 1, "",
                wedgeKeys.Length + 1, 1);

            table.SetColumnWidth(0, drawingData.TablePositions["dimension"].GetValues(Unit.Meter)[2],
                (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
            table.GridLineWeight = (int)swLineWeights_e.swLW_NONE;

            table.SetHeader((int)swTableHeaderPosition_e.swTableHeader_Top, 1);
            table.Title = "Dimensions";
            table.TitleVisible = true;

            table.set_Text(0, 0, header);

            for (int i = 0; i < wedgeKeys.Length; i++)
            {
                string key = wedgeKeys[i];
                string valMM = wedgeDimensions[key].GetValue(Unit.Millimeter).ToString("F2") + " mm";
                string tolMM = $"+{wedgeDimensions[key].GetTolerance(Unit.Millimeter, "+").ToString("F2")}/-{wedgeDimensions[key].GetTolerance(Unit.Millimeter, "-").ToString("F2")}";
                table.set_Text(i + 1, 0, $"{key}: {valMM} ({tolMM})");
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool CreateLabelAsTable(DataStorage position, DrawingData drawingData)
    {
        try
        {
            if (drawingData.LabelAsItems.Length == 0)
                return false;

            var table = (TableAnnotation)_swDrawing.InsertTableAnnotation2(false, position.GetValues(Unit.Meter)[0], position.GetValues(Unit.Meter)[1], 1, "",
                drawingData.LabelAsItems.Length, 1);

            table.SetColumnWidth(0, drawingData.TablePositions["label_as"].GetValues(Unit.Meter)[2],
                (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
            table.GridLineWeight = (int)swLineWeights_e.swLW_NONE;
            table.Title = "Label As";
            table.TitleVisible = true;

            for (int i = 0; i < drawingData.LabelAsItems.Length; i++)
            {
                table.set_Text(i, 0, drawingData.LabelAsItems[i]);
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool CreatePolishTable(DataStorage position, DrawingData drawingData)
    {
        try
        {
            if (drawingData.PolishItems.Length == 0)
                return false;

            var table = (TableAnnotation)_swDrawing.InsertTableAnnotation2(false, position.GetValues(Unit.Meter)[0], position.GetValues(Unit.Meter)[1], 1, "",
                drawingData.PolishItems.Length, 1);

            table.SetColumnWidth(0, drawingData.TablePositions["polish"].GetValues(Unit.Meter)[2],
                (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
            table.GridLineWeight = (int)swLineWeights_e.swLW_NONE;
            table.Title = "Polish";
            table.TitleVisible = true;

            for (int i = 0; i < drawingData.PolishItems.Length; i++)
            {
                table.set_Text(i, 0, drawingData.PolishItems[i]);
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool CreateHowToOrderTable(DataStorage position, string header, DrawingData drawingData)
    {
        try
        {
            int rowCount = drawingData.HowToOrderInfo.ContainsKey("packaging") && !string.IsNullOrWhiteSpace(drawingData.HowToOrderInfo["packaging"])
                ? 3 : 2;

            var table = (TableAnnotation)_swDrawing.InsertTableAnnotation2(false, position.GetValues(Unit.Meter)[0], position.GetValues(Unit.Meter)[1], 1, "",
                rowCount, 1);

            table.SetColumnWidth(0, drawingData.TablePositions["how_to_order"].GetValues(Unit.Meter)[2],
                (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
            table.GridLineWeight = (int)swLineWeights_e.swLW_NONE;
            table.Title = "How To Order";
            table.TitleVisible = true;

            table.set_Text(0, 0, header);
            table.set_Text(1, 0, drawingData.Title);

            if (rowCount == 3)
                table.set_Text(2, 0, drawingData.HowToOrderInfo["packaging"]);

            return true;
        }
        catch
        {
            return false;
        }
    }
}
