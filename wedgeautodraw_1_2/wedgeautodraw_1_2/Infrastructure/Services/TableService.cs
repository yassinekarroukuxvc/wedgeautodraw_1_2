using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

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

    private TableAnnotation InsertBasicTable(DataStorage position, int rows, string title, double width)
    {
        try
        {
            var table = (TableAnnotation)_swDrawing.InsertTableAnnotation2(
                false,
                position.GetValues(Unit.Meter)[0],
                position.GetValues(Unit.Meter)[1],
                1,
                "",
                rows,
                1);

            if (table == null)
            {
                Logger.Warn($"Failed to insert table '{title}'.");
                return null;
            }

            table.SetColumnWidth(0, width, (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
            table.GridLineWeight = (int)swLineWeights_e.swLW_NONE;
            table.Title = title;
            table.TitleVisible = true;

            Logger.Info($"Table '{title}' inserted successfully.");
            return table;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error inserting table '{title}': {ex.Message}");
            return null;
        }
    }

    public bool CreateDimensionTable(DataStorage position, string[] wedgeKeys, string header, DrawingData drawingData, NamedDimensionValues wedgeDimensions)
    {
        var validRows = new List<string>();

        foreach (var key in wedgeKeys)
        {
            if (wedgeDimensions.TryGet(key, out var dataStorage) && dataStorage != null)
            {
                double valueInch = dataStorage.GetValue(Unit.Inch);
                double upperTolInch = dataStorage.GetTolerance(Unit.Inch, "+");
                double lowerTolInch = dataStorage.GetTolerance(Unit.Inch, "-");

                double valueMm = dataStorage.GetValue(Unit.Millimeter);
                double upperTolMm = dataStorage.GetTolerance(Unit.Millimeter, "+");
                double lowerTolMm = dataStorage.GetTolerance(Unit.Millimeter, "-");

                if (!double.IsNaN(valueInch))
                {
                    bool isRef = IsRef(upperTolInch, lowerTolInch);

                    string tolStrInch = isRef ? "" : FormatTolerance(upperTolInch, lowerTolInch, 4, true);
                    string tolStrMm = isRef ? "" : FormatTolerance(upperTolMm, lowerTolMm, 4, false);

                    // Format inch value without leading 0
                    string inchStr = TrimLeadingZero(valueInch.ToString("F4"));

                    string valStr = $"{key} = {inchStr} {tolStrInch} [{valueMm:F3} {tolStrMm}]";

                    if (isRef)
                        valStr += " (REF)";

                    validRows.Add(valStr.Trim());
                }
            }
        }

        if (validRows.Count == 0)
        {
            Logger.Warn("No valid dimension entries found. Skipping dimension table creation.");
            return false;
        }

        var table = InsertBasicTable(position, validRows.Count + 1, "Dimensions", drawingData.TablePositions[Constants.DimensionTable].GetValues(Unit.Meter)[2]);
        if (table == null) return false;

        table.SetHeader((int)swTableHeaderPosition_e.swTableHeader_Top, 1);
        table.set_Text(0, 0, header);

        for (int i = 0; i < validRows.Count; i++)
        {
            table.set_Text(i + 1, 0, validRows[i]);
        }

        return true;
    }

    private string FormatTolerance(double upper, double lower, int precision, bool removeLeadingZero)
    {
        bool upperValid = !double.IsNaN(upper);
        bool lowerValid = !double.IsNaN(lower);

        string Format(double val) =>
            Math.Abs(val) < 1e-6 ? "0" :
            removeLeadingZero ? TrimLeadingZero(val.ToString($"F{precision}")) :
            val.ToString($"F{precision}");

        if (!upperValid && !lowerValid)
            return "";

        bool upperZero = Math.Abs(upper) < 1e-6;
        bool lowerZero = Math.Abs(lower) < 1e-6;

        if (!upperZero && !lowerZero && Math.Abs(upper - lower) < 1e-6)
            return $"±{Format(upper)}";

        if (!upperZero && lowerZero)
            return $"+{Format(upper)}/-0";

        if (upperZero && !lowerZero)
            return $"+0/-{Format(lower)}";

        if (!upperZero && !lowerZero)
            return $"+{Format(upper)}/-{Format(lower)}";

        return "";
    }

    private bool IsRef(double upper, double lower)
    {
        return (double.IsNaN(upper) && double.IsNaN(lower)) ||
               (Math.Abs(upper) < 1e-6 && Math.Abs(lower) < 1e-6);
    }

    private string TrimLeadingZero(string input)
    {
        return input.StartsWith("0.") ? input.Substring(1) : input;
    }


    public bool CreateLabelAsTable(DataStorage position, DrawingData drawingData)
    {
        if (drawingData.LabelAsItems.Length == 0)
        {
            Logger.Warn("No Label As items to create table.");
            return false;
        }

        var table = InsertBasicTable(position, drawingData.LabelAsItems.Length, "Label As", drawingData.TablePositions[Constants.LabelAsTable].GetValues(Unit.Meter)[2]);
        if (table == null) return false;

        for (int i = 0; i < drawingData.LabelAsItems.Length; i++)
        {
            table.set_Text(i, 0, drawingData.LabelAsItems[i]);
        }

        return true;
    }

    public bool CreatePolishTable(DataStorage position, DrawingData drawingData)
    {
        if (drawingData.PolishItems.Length == 0)
        {
            Logger.Warn("No Polish items to create table.");
            return false;
        }

        var table = InsertBasicTable(position, drawingData.PolishItems.Length, "Polish", drawingData.TablePositions[Constants.PolishTable].GetValues(Unit.Meter)[2]);
        if (table == null) return false;

        for (int i = 0; i < drawingData.PolishItems.Length; i++)
        {
            table.set_Text(i, 0, drawingData.PolishItems[i]);
        }

        return true;
    }

    public bool CreateHowToOrderTable(DataStorage position, string header, DrawingData drawingData)
    {
        int rowCount = drawingData.HowToOrderInfo.ContainsKey("packaging") && !string.IsNullOrWhiteSpace(drawingData.HowToOrderInfo["packaging"])
            ? 3 : 2;

        var table = InsertBasicTable(position, rowCount, "How To Order", drawingData.TablePositions[Constants.HowToOrderTable].GetValues(Unit.Meter)[2]);
        if (table == null) return false;

        table.set_Text(0, 0, header);
        table.set_Text(1, 0, drawingData.Title);

        if (rowCount == 3)
            table.set_Text(2, 0, drawingData.HowToOrderInfo["packaging"]);

        return true;
    }
    public bool CreateDimensionNote(DataStorage position, string[] wedgeKeys, string header, DrawingData drawingData, NamedDimensionValues wedgeDimensions)
    {
        try
        {
            var validLines = new List<string>();

            foreach (var key in wedgeKeys)
            {
                if (wedgeDimensions.TryGet(key, out var dataStorage) && dataStorage != null)
                {
                    double valueInch = dataStorage.GetValue(Unit.Inch);
                    double upperTolInch = dataStorage.GetTolerance(Unit.Inch, "+");
                    double lowerTolInch = dataStorage.GetTolerance(Unit.Inch, "-");

                    double valueMm = dataStorage.GetValue(Unit.Millimeter);
                    double upperTolMm = dataStorage.GetTolerance(Unit.Millimeter, "+");
                    double lowerTolMm = dataStorage.GetTolerance(Unit.Millimeter, "-");

                    if (!double.IsNaN(valueInch))
                    {
                        bool isRef = IsRef(upperTolInch, lowerTolInch);

                        string tolStrInch = isRef ? "" : FormatTolerance(upperTolInch, lowerTolInch, 4, true);
                        string tolStrMm = isRef ? "" : FormatTolerance(upperTolMm, lowerTolMm, 4, false);

                        string inchStr = TrimLeadingZero(valueInch.ToString("F4"));

                        string valStr = $"{key} = {inchStr} {tolStrInch} [{valueMm:F3} {tolStrMm}]";

                        if (isRef)
                            valStr += " (REF)";

                        validLines.Add(valStr.Trim());
                    }
                }
            }

            if (validLines.Count == 0)
            {
                Logger.Warn("No valid dimension entries found. Skipping dimension note creation.");
                return false;
            }

            // Build note text
            var noteText = $"{header}\n";
            foreach (var line in validLines)
            {
                noteText += $"{line}\n";
            }

            double[] pos = position.GetValues(Unit.Meter);
            double posX = pos[0];
            double posY = pos[1];
            double textHeight = 0.004;  // Example text height in meters (~ 4 mm)

            // Now use correct method signature:
            var noteObj = _swDrawing.CreateText2(noteText, textHeight, posX, posY, 0.0, 0.0);

            if (noteObj == null)
            {
                Logger.Warn("Failed to insert dimension note.");
                return false;
            }

            Logger.Success("Dimension note created successfully.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error creating dimension note: {ex.Message}");
            return false;
        }
    }

}
