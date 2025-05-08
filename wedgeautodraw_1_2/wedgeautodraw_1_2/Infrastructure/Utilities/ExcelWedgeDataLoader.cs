using ClosedXML.Excel;
using System.Globalization;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Utilities;

public class ExcelWedgeDataLoader
{
    private readonly string _excelFilePath;

    public ExcelWedgeDataLoader(string excelFilePath)
    {
        _excelFilePath = excelFilePath;
    }

    public List<(WedgeData Wedge, DrawingData Drawing)> LoadAllEntries()
    {
        var result = new List<(WedgeData, DrawingData)>();

        if (!File.Exists(_excelFilePath))
        {
            Console.WriteLine($"Excel file not found: {_excelFilePath}");
            return result;
        }

        using var workbook = new XLWorkbook(_excelFilePath);
        var worksheet = workbook.Worksheet(1);
        var rows = worksheet.RangeUsed().RowsUsed().Skip(1).ToList(); // Skip header row
        var header = worksheet.Row(1);

        var columnMap = new Dictionary<string, int>();
        for (int c = 1; c <= header.CellCount(); c++)
        {
            string key = header.Cell(c).GetValue<string>().Trim();
            if (!string.IsNullOrEmpty(key)) columnMap[key] = c;
        }

        // Detect available dimensions
        var dimensionKeys = columnMap.Keys
            .Where(k => k.EndsWith("_NOM"))
            .Select(k => k[..^4])
            .Distinct()
            .ToList();

        // Parallel processing
        var parallelResult = rows.AsParallel().Select(row =>
        {
            var wedge = BuildWedge(row, columnMap, dimensionKeys);
            var drawing = BuildDrawing(row, columnMap, wedge);
            return (wedge, drawing);
        }).ToList();

        return parallelResult;
    }

    private WedgeData BuildWedge(IXLRangeRow row, Dictionary<string, int> map, List<string> dimensionKeys)
    {
        var wedge = new WedgeData
        {
            EngravedText = GetCell(row, map, "wedge_title")
        };

        wedge.Metadata["drawing_number"] = GetCell(row, map, "drawing#");
        wedge.Metadata["drawing_title"] = GetCell(row, map, "drawing#") + "-DW";
        wedge.Metadata["wedge_title"] = wedge.EngravedText;
        wedge.EngravedText = wedge.EngravedText.Replace("¶", " ");

        wedge.Metadata["Source"] = _excelFilePath;

        foreach (var key in dimensionKeys)
        {
            string nom = GetCell(row, map, key + "_NOM");
            string upper = GetCell(row, map, key + "_UTOL");
            string lower = GetCell(row, map, key + "_LTOL");

            if (!string.IsNullOrWhiteSpace(nom))
            {
                if (!IsAngle(key))
                {
                    nom = ConvertInchToMillimeter(nom);
                    upper = ConvertInchToMillimeter(upper);
                    lower = ConvertInchToMillimeter(lower);
                }

                var data = new DataStorage(nom, upper, lower);
                data.SetUnit(IsAngle(key) ? Unit.Degree : Unit.Millimeter);
                wedge.Dimensions[key] = data;
            }
        }

        if (!wedge.Dimensions.ContainsKey("SymmetryTolerance"))
        {
            wedge.Dimensions["SymmetryTolerance"] = new DataStorage(0.04);
            wedge.Dimensions["SymmetryTolerance"].SetUnit(Unit.Millimeter);
        }

        return wedge;
    }

    private DrawingData BuildDrawing(IXLRangeRow row, Dictionary<string, int> map, WedgeData wedge)
    {
        var drawing = new DrawingData();

        string drawingNumber = GetCell(row, map, "drawing#");
        drawing.TitleInfo["number"] = drawingNumber;
        drawing.TitleInfo["info"] = GetCell(row, map, "drawing_comments");

        drawing.Title = drawingNumber + " - Production Copy";
        drawing.TitleBlockInfo["DRAWING_NUMBER"] = drawingNumber + "-DW";
        drawing.TitleBlockInfo["TITLE"] = drawing.Title;
        drawing.TitleBlockInfo["COMPANY_NAME"] = "SMALL PRECISION TOOLS";
        drawing.TitleBlockInfo["DRAWN_BY"] = "AUTODRAW SERVICE";
        drawing.TitleBlockInfo["DRAWN_ON"] = DateTime.Now.ToString("MM/dd/yyyy");
        Logger.Success(drawing.TitleBlockInfo["DRAWN_ON"]);

        drawing.HowToOrderInfo["number"] = drawingNumber;
        drawing.HowToOrderInfo["packaging"] = GetCell(row, map, "packaging");

        drawing.LabelAsItems = GetCell(row, map, "engrave").Split('¶');
        drawing.PolishItems = GetCell(row, map, "finishing").Split('¶');

        drawing.DrawingType = DrawingType.Production;

        return drawing;
    }

    private string GetCell(IXLRangeRow row, Dictionary<string, int> map, string columnName)
    {
        if (map.TryGetValue(columnName, out int colIndex))
            return row.Cell(colIndex).GetValue<string>().Trim();
        return string.Empty;
    }

    private string ConvertInchToMillimeter(string valueInInch)
    {
        if (double.TryParse(valueInInch, NumberStyles.Float, CultureInfo.InvariantCulture, out double inch))
        {
            double mm = inch * 25.4;
            return mm.ToString(CultureInfo.InvariantCulture);
        }
        return valueInInch;
    }

    private bool IsAngle(string key)
    {
        return key is "ISA" or "FA" or "BA" or "GA" or "FL_groove_angle";
    }
}
