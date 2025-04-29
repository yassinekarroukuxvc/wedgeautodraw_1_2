using ClosedXML.Excel;
using System.Globalization;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Infrastructure.Utilities;

public class ExcelWedgeDataLoader
{
    private readonly string _excelFilePath;

    public ExcelWedgeDataLoader(string excelFilePath)
    {
        _excelFilePath = excelFilePath;
    }

    public List<WedgeData> LoadWedgeDataList()
    {
        var wedgeList = new List<WedgeData>();

        if (!File.Exists(_excelFilePath))
        {
            Console.WriteLine($"Excel file not found: {_excelFilePath}");
            return wedgeList;
        }

        using var workbook = new XLWorkbook(_excelFilePath);
        var worksheet = workbook.Worksheet(1);
        var rows = worksheet.RangeUsed().RowsUsed();

        var header = worksheet.Row(1);
        var columnMap = new Dictionary<string, int>();
        for (int c = 1; c <= header.CellCount(); c++)
        {
            string key = header.Cell(c).GetValue<string>().Trim();
            if (!string.IsNullOrEmpty(key)) columnMap[key] = c;
        }

        // Dynamically detect available dimensions
        var dimensionKeys = columnMap.Keys
            .Where(k => k.EndsWith("_NOM"))
            .Select(k => k[..^4])
            .Distinct()
            .ToList();

        foreach (var row in rows.Skip(1))
        {
            var wedge = new WedgeData();

            wedge.Metadata["drawing_number"] = GetCell(row, columnMap, "drawing#");
            wedge.Metadata["drawing_title"] = GetCell(row, columnMap, "drawing_title");
            wedge.Metadata["wedge_title"] = GetCell(row, columnMap, "wedge_title");
            wedge.EngravedText = GetCell(row, columnMap, "wedge_title");

            foreach (var key in dimensionKeys)
            {
                string nom = GetCell(row, columnMap, key + "_NOM");
                string upper = GetCell(row, columnMap, key + "_UTOL");
                string lower = GetCell(row, columnMap, key + "_LTOL");

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

            wedge.Metadata["Source"] = _excelFilePath;
            wedgeList.Add(wedge);
        }

        return wedgeList;
    }

    public List<DrawingData> LoadDrawingDataList()
    {
        var drawingList = new List<DrawingData>();

        if (!File.Exists(_excelFilePath))
        {
            Console.WriteLine($"Excel file not found: {_excelFilePath}");
            return drawingList;
        }

        using var workbook = new XLWorkbook(_excelFilePath);
        var worksheet = workbook.Worksheet(1);
        var rows = worksheet.RangeUsed().RowsUsed();

        var header = worksheet.Row(1);
        var columnMap = new Dictionary<string, int>();
        for (int c = 1; c <= header.CellCount(); c++)
        {
            string key = header.Cell(c).GetValue<string>().Trim();
            if (!string.IsNullOrEmpty(key)) columnMap[key] = c;
        }

        foreach (var row in rows.Skip(1))
        {
            var drawing = new DrawingData();

            drawing.TitleInfo["number"] = GetCell(row, columnMap, "drawing#");
            drawing.TitleInfo["info"] = GetCell(row, columnMap, "drawing_comments");
            drawing.Title = drawing.TitleInfo["number"] + " - Production Copy";
            drawing.TitleBlockInfo["DRAWING_NUMBER"] = drawing.TitleInfo["number"] + "-DW";
            drawing.TitleBlockInfo["TITLE"] = drawing.Title;
            drawing.TitleBlockInfo["COMPANY_NAME"] = "SMALL PRECISION TOOLS";
            drawing.TitleBlockInfo["DRAWN_BY"] = "AUTODRAW SERVICE";
            drawing.TitleBlockInfo["DRAWN_ON"] = DateTime.Now.ToString("yyyy-MM-dd");

            drawing.HowToOrderInfo["number"] = drawing.TitleInfo["number"];
            drawing.HowToOrderInfo["packaging"] = GetCell(row, columnMap, "packaging");

            drawing.LabelAsItems = GetCell(row, columnMap, "engrave").Split('¶');
            drawing.PolishItems = GetCell(row, columnMap, "finishing").Split('¶');

            drawing.DrawingType = DrawingType.Production;
            drawingList.Add(drawing);
        }

        return drawingList;
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
