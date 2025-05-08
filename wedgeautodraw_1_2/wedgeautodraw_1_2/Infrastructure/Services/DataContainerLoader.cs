using System.Globalization;
using System.Text.RegularExpressions;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services;

public class DataContainerLoader : IDataContainerLoader
{
    private readonly string _equationFilePath;
    private string _toleranceFilePath;

    public DataContainerLoader(string equationFilePath)
    {
        _equationFilePath = equationFilePath;
    }

    public WedgeData LoadWedgeData(string toleranceFilePath)
    {
        var wedgeData = new WedgeData();

        if (!File.Exists(_equationFilePath))
        {
            Console.WriteLine($"Equation file not found: {_equationFilePath}");
            return wedgeData;
        }

        var lines = File.ReadAllLines(_equationFilePath);
        var regex = new Regex("\"(?<key>[^\"]+)\"=\\s*(?<value>[-+]?\\d*\\.?\\d+)(?<unit>mm|in|deg|rad|m)");

        foreach (var line in lines)
        {
            var match = regex.Match(line);
            if (match.Success)
            {
                string key = match.Groups["key"].Value;
                string valueStr = match.Groups["value"].Value;
                string unitStr = match.Groups["unit"].Value;

                if (double.TryParse(valueStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
                {
                    var data = new DataStorage(value);

                    if (Enum.TryParse(unitStr, true, out Unit parsedUnit))
                    {
                        data.SetUnit(parsedUnit);
                    }

                    wedgeData.Dimensions[key] = data;
                }
            }
        }

        // Apply tolerances from tolerance file (only upper/lower tol)
        if (!string.IsNullOrEmpty(_toleranceFilePath) && File.Exists(_toleranceFilePath))
        {
            var tolRegex = new Regex("\"(?<key>[^\"]+)\"\\s*\\+(?<upper>[-+]?\\d*\\.?\\d+)\\s*-?(?<lower>[-+]?\\d*\\.?\\d+)\\s*(?<unit>mm|in|deg|rad|m)?");

            var tolLines = File.ReadAllLines(_toleranceFilePath);
            foreach (var line in tolLines)
            {
                var match = tolRegex.Match(line);
                if (match.Success)
                {
                    string key = match.Groups["key"].Value;
                    string upper = match.Groups["upper"].Value;
                    string lower = match.Groups["lower"].Value;

                    if (wedgeData.Dimensions.GetAll().ContainsKey(key))
                    {
                        var original = wedgeData.Dimensions[key];
                        wedgeData.Dimensions[key] = new DataStorage(
                            original.GetValue(Unit.Millimeter).ToString(CultureInfo.InvariantCulture),
                            upper,
                            lower
                        );
                    }
                }
            }
        }

        wedgeData.EngravedText = "Sample";
        wedgeData.Metadata["Source"] = _equationFilePath;
        wedgeData.Dimensions["SymmetryTolerance"] = new DataStorage(0.04); // value in mm
        wedgeData.Dimensions["SymmetryTolerance"].SetUnit(Unit.Millimeter);

        return wedgeData;
    }

    public DrawingData LoadDrawingData(WedgeData wedgeData, string configFilePath)
    {
        var drawingData = new DrawingData();
        if (!File.Exists(configFilePath))
        {
            Logger.Warn("Configuration file not found.");
            return drawingData;
        }

        var config = File.ReadAllLines(configFilePath)
            .Where(line => !string.IsNullOrWhiteSpace(line) && !line.Trim().StartsWith("#"))
            .Select(line => line.Split('=')).ToDictionary(parts => parts[0].Trim(), parts => parts[1].Trim());

        double Get(string key) => config.TryGetValue(key, out var v) && double.TryParse(v, out var d) ? d : double.NaN;
        string GetStr(string key) => config.TryGetValue(key, out var v) ? v : string.Empty;

        // View Scales
        double defaultScale = Get("scaling_dsv");
        double W_value = wedgeData.Dimensions.ContainsKey("W") ? wedgeData.Dimensions["W"].GetValue(Unit.Millimeter) : 10.0;

        // Decrease scale only if W is above a threshold
        double adjustedScale = W_value >= 0.7 ? Math.Max(defaultScale * (1.0 / W_value), 0.2) : defaultScale;
        adjustedScale = Math.Round(adjustedScale, 3);
        Console.WriteLine($"The Default Scale {defaultScale} /// The Adjusted Scale {adjustedScale} ");
        drawingData.ViewScales["Front_view"] = new DataStorage(Get("scaling_fsv"));
        drawingData.ViewScales["Side_view"] = new DataStorage(Get("scaling_fsv"));
        drawingData.ViewScales["Top_view"] = new DataStorage(Get("scaling_fsv"));
        drawingData.ViewScales["Detail_view"] = new DataStorage(adjustedScale);
        drawingData.ViewScales["Section_view"] = new DataStorage(adjustedScale);

        double fx = 0.0;
        double td = wedgeData.Dimensions["TD"].GetValue(Unit.Millimeter);
        double tdf = wedgeData.Dimensions["TDF"].GetValue(Unit.Millimeter);
        double fl = wedgeData.Dimensions["FL"].GetValue(Unit.Millimeter);

        double sectionViewCenterX = fx == 0.0 || double.IsNaN(fx)
            ? Get("section_view_posX") + drawingData.ViewScales["Section_view"].GetValue(Unit.Millimeter) * (td - tdf) / 2
            : Get("section_view_posX") + drawingData.ViewScales["Section_view"].GetValue(Unit.Millimeter) * ((tdf - fl) / 2 - fx);

        // View Positions
        drawingData.ViewPositions["Front_view"] = new DataStorage(new[] { Get("front_view_posX"), Get("front_view_posY") });
        drawingData.ViewPositions["Side_view"] = new DataStorage(new[] {
    Get("front_view_posX") + Get("side_view_dX"), Get("front_view_posY") });
        drawingData.ViewPositions["Top_view"] = new DataStorage(new[] {
    Get("front_view_posX") + Get("side_view_dX"), Get("front_view_posY") + Get("top_view_dY") });
        drawingData.ViewPositions["Detail_view"] = new DataStorage(new[] { Get("detail_view_posX"), Get("detail_view_posY") });
        drawingData.ViewPositions["Section_view"] = new DataStorage(new[] {
    Get("detail_view_posX") + sectionViewCenterX, Get("detail_view_posY") });


        // Table Positions
        drawingData.TablePositions["dimension"] = new DataStorage(new[] {
    Get("dim_table_posX"), Get("dim_table_posY"), Get("dim_table_width") });
        drawingData.TablePositions["how_to_order"] = new DataStorage(new[] {
    Get("how_to_order_posX"), Get("how_to_order_posY"), Get("how_to_order_width") });
        drawingData.TablePositions["label_as"] = new DataStorage(new[] {
    Get("label_as_posX"), Get("label_as_posY"), Get("label_as_width") });
        drawingData.TablePositions["polish"] = new DataStorage(new[] {
    Get("polish_posX"), Get("polish_posY"), Get("polish_width") });

        // Breakline Data
        drawingData.BreaklineData["Front_viewLowerPartLength"] = new DataStorage(Get("length_lower_section_fsv"));
        drawingData.BreaklineData["Front_viewUpperPartLength"] = new DataStorage(Get("length_upper_section_fsv"));
        drawingData.BreaklineData["Front_viewBreaklineGap"] = new DataStorage(Get("breakline_gap_fsv"));
        drawingData.BreaklineData["Side_viewLowerPartLength"] = new DataStorage(Get("length_lower_section_fsv"));
        drawingData.BreaklineData["Side_viewUpperPartLength"] = new DataStorage(Get("length_upper_section_fsv"));
        drawingData.BreaklineData["Side_viewBreaklineGap"] = new DataStorage(Get("breakline_gap_fsv"));
        drawingData.BreaklineData["Detail_viewLowerPartLength"] = new DataStorage(Get("length_lower_section_dsv"));
        drawingData.BreaklineData["Detail_viewUpperPartLength"] = new DataStorage(0);
        drawingData.BreaklineData["Detail_viewBreaklineGap"] = new DataStorage(Get("breakline_gap_dsv"));
        drawingData.BreaklineData["Section_viewLowerPartLength"] = new DataStorage(Get("length_lower_section_dsv"));
        drawingData.BreaklineData["Section_viewUpperPartLength"] = new DataStorage(0);
        drawingData.BreaklineData["Section_viewBreaklineGap"] = new DataStorage(Get("breakline_gap_dsv"));

        // Title Info
        drawingData.TitleInfo["info"] = GetStr("how_to_order_info");
        drawingData.TitleInfo["number"] = (string)wedgeData.Metadata["drawing_number"];
        drawingData.Title = wedgeData.Metadata["drawing_title"] + " - Production Copy";

        // Title Block Info
        drawingData.TitleBlockInfo["Material"] = GetStr("material");
        drawingData.TitleBlockInfo["Autor"] = GetStr("author");
        drawingData.TitleBlockInfo["DRAWN_BY"] = "AUTODRAW SERVICE";
        //drawingData.TitleBlockInfo["DRAWN_ON"] = "NOT SET YET!";
        drawingData.TitleBlockInfo["COMPANY_NAME"] = "SMALL PRECISION TOOLS";
        drawingData.TitleBlockInfo["TITLE"] = drawingData.Title;
        drawingData.TitleBlockInfo["DRAWING_NUMBER"] = drawingData.TitleInfo["number"] + "-DW";
        drawingData.TitleBlockInfo["ADDRESS"] = "1330 CLEGG STREET PETALUMA, CALIFORNIA 94954";
        drawingData.TitleBlockInfo["TYPE"] = "PRODUCTION COPY";
        drawingData.TitleBlockInfo["SCALING_FRONT_SIDE_TOP_VIEW"] = Get("scaling_fsv").ToString();
        drawingData.TitleBlockInfo["SCALING_DETAIL_SECTION_VIEW"] = Get("scaling_dsv").ToString();
        drawingData.TitleBlockInfo["DRAWN_ON"] = DateTime.Now.ToString("MM-dd-yy");


        // How To Order Info
        drawingData.HowToOrderInfo["number"] = drawingData.TitleInfo["number"];
        drawingData.HowToOrderInfo["packaging"] = GetStr("packaging");

        // Drawing Type
        drawingData.DrawingType = DrawingType.Production;

        // Label As and Polish Items
        drawingData.LabelAsItems = GetStr("engrave").Split('¶');
        drawingData.PolishItems = GetStr("polish_text").Split('¶');

        // Dimension Keys in Table
        drawingData.DimensionKeysInTable = GetStr("dimension_keys_in_table").Split(',');

        // Dimension Styles
        DynamicDimensionStyler.ApplyDynamicStyles(drawingData, wedgeData);

        return drawingData;
    }

}
