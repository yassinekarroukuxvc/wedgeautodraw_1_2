using ClosedXML.Excel;
using System.Globalization;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Utilities;

public class ExcelWedgeDataLoader
{
    private readonly string _excelFilePath;
    private readonly WedgeType _wedgeType;

    public ExcelWedgeDataLoader(string excelFilePath, WedgeType wedgeType)
    {
        _excelFilePath = excelFilePath;
        _wedgeType = wedgeType;
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
        var rows = worksheet.RangeUsed().RowsUsed().Skip(1).ToList();
        var header = worksheet.Row(1);

        var columnMap = new Dictionary<string, int>();
        for (int c = 1; c <= header.CellCount(); c++)
        {
            string key = header.Cell(c).GetValue<string>().Trim();
            if (!string.IsNullOrEmpty(key)) columnMap[key] = c;
        }

        var dimensionKeys = columnMap.Keys
            .Where(k => k.EndsWith("_NOM"))
            .Select(k => k[..^4])
            .Distinct()
            .ToList();

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
            EngravedText = GetCell(row, map, "wedge_title"),
            WedgeType = _wedgeType // ✅ Set from constructor
        };

        wedge.Metadata["drawing_number"] = GetCell(row, map, "drawing#");
        wedge.Metadata["drawing_title"] = wedge.Metadata["drawing_number"] + "-DW";
        wedge.Metadata["wedge_title"] = wedge.EngravedText.Replace("¶", " ");

        wedge.OverlayCalibration = string.Empty;
        wedge.OverlayScaling = 1.0;

        string rawOverlay = GetCell(row, map, "overlay_calibration");
        wedge.Coining = GetCell(row, map, "coining").Replace("¶", " ");

        if (!string.IsNullOrWhiteSpace(rawOverlay))
        {
            try
            {
                string[] parts = rawOverlay.Split('|');
                if (parts.Length == 2)
                {
                    string scalingPart = parts[0].Trim();
                    if (scalingPart.EndsWith("X", StringComparison.OrdinalIgnoreCase))
                    {
                        string numberPart = scalingPart.Replace("X", "").Trim();
                        if (double.TryParse(numberPart, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsedScaling))
                        {
                            wedge.OverlayScaling = parsedScaling;
                        }
                    }

                    string calibPart = parts[1].Trim();
                    if (calibPart.EndsWith("um", StringComparison.OrdinalIgnoreCase))
                    {
                        string numberPart = calibPart.Replace("um", "").Trim();
                        wedge.OverlayCalibration = numberPart;
                    }
                }
                else
                {
                    Logger.Warn($"overlay_calibration format unexpected: '{rawOverlay}'");
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"Failed to parse overlay_calibration: '{rawOverlay}' → {ex.Message}");
            }
        }

        wedge.Metadata["Source"] = _excelFilePath;

        foreach (var key in dimensionKeys)
        {
            string nom = GetCell(row, map, key + "_NOM");
            string upper = GetCell(row, map, key + "_UTOL");
            string lower = GetCell(row, map, key + "_LTOL");

            if (string.IsNullOrWhiteSpace(nom)) nom = "0";
            if (string.IsNullOrWhiteSpace(upper)) upper = "0";
            if (string.IsNullOrWhiteSpace(lower)) lower = "0";

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
        string wedgeTitle = GetCell(row, map, "wedge_title");

        drawing.Title = drawingNumber + " - Production Copy";
        drawing.TitleBlockInfo["DRAWING_NUMBER"] = drawingNumber + "-DW";
        drawing.TitleBlockInfo["TITLE"] = wedgeTitle;
        drawing.TitleBlockInfo["COMPANY_NAME"] = "SMALL PRECISION TOOLS";
        drawing.TitleBlockInfo["DRAWN_BY"] = "AUTODRAW SERVICE";
        drawing.TitleBlockInfo["DRAWN_ON"] = DateTime.Now.ToString("MM/dd/yyyy");

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
