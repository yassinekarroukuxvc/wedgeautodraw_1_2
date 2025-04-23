using System.Globalization;
using System.Text.RegularExpressions;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;

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
            Console.WriteLine("Configuration file not found.");
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
        Console.WriteLine("FrontViewBreaklineGap " + drawingData.BreaklineData["Front_viewBreaklineGap"].GetValue(Unit.Millimeter));
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
        drawingData.TitleInfo["number"] = GetStr("how_to_order_number");
        drawingData.Title = drawingData.TitleInfo["number"] + " - Production Copy";

        // Title Block Info
        drawingData.TitleBlockInfo["Material"] = GetStr("material");
        drawingData.TitleBlockInfo["Autor"] = GetStr("author");
        drawingData.TitleBlockInfo["DRAWN_BY"] = "AUTODRAW SERVICE";
        drawingData.TitleBlockInfo["DRAWN_ON"] = "NOT SET YET!";
        drawingData.TitleBlockInfo["COMPANY_NAME"] = "SMALL PRECISION TOOLS";
        drawingData.TitleBlockInfo["TITLE"] = drawingData.Title;
        drawingData.TitleBlockInfo["DRAWING_NUMBER"] = drawingData.TitleInfo["number"] + "-DW";
        drawingData.TitleBlockInfo["ADDRESS"] = "1330 CLEGG STREET PETALUMA, CALIFORNIA 94954";
        drawingData.TitleBlockInfo["TYPE"] = "PRODUCTION COPY";
        drawingData.TitleBlockInfo["SCALING_FRONT_SIDE_TOP_VIEW"] = Get("scaling_fsv").ToString();
        drawingData.TitleBlockInfo["SCALING_DETAIL_SECTION_VIEW"] = Get("scaling_dsv").ToString();

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

        // === DimensionStyles ===
        var fsv = drawingData.ViewScales["Front_view"].GetValue(Unit.Millimeter);
        var dsv = drawingData.ViewScales["Detail_view"].GetValue(Unit.Millimeter);
        var tsv = drawingData.ViewScales["Top_view"].GetValue(Unit.Millimeter);
        var ssv = drawingData.ViewScales["Side_view"].GetValue(Unit.Millimeter);
        var secv = drawingData.ViewScales["Section_view"].GetValue(Unit.Millimeter);

        var front = drawingData.ViewPositions["Front_view"].GetValues(Unit.Millimeter);
        var top = drawingData.ViewPositions["Top_view"].GetValues(Unit.Millimeter);
        var side = drawingData.ViewPositions["Side_view"].GetValues(Unit.Millimeter);
        var detail = drawingData.ViewPositions["Detail_view"].GetValues(Unit.Millimeter);
        var section = drawingData.ViewPositions["Section_view"].GetValues(Unit.Millimeter);

        var W = wedgeData.Dimensions["W"].GetValue(Unit.Millimeter);
        var GD = wedgeData.Dimensions["GD"].GetValue(Unit.Millimeter);
        var TD = wedgeData.Dimensions["TD"].GetValue(Unit.Millimeter);
        var TDF = wedgeData.Dimensions["TDF"].GetValue(Unit.Millimeter);
        var FL = wedgeData.Dimensions["FL"].GetValue(Unit.Millimeter);

        var detailLowerLength = drawingData.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);

        // Begin dimensioning setup
        drawingData.DimensionStyles["TL"] = new DimensioningStorage(new DataStorage(new[] {
            front[0] - fsv * TD / 2 - 7.5, front[1]
        }));
        drawingData.DimensionStyles["EngravingStart"] = new DimensioningStorage(new DataStorage(new[] {
            front[0] + fsv * TD / 2 + 4, front[1] + 45
        }));
        drawingData.DimensionStyles["TDF"] = new DimensioningStorage(new DataStorage(new[] {
            top[0] + tsv * TDF / 2 + 20, top[1] + tsv * TD / 2 + 3
        }));
        drawingData.DimensionStyles["TD"] = new DimensioningStorage(new DataStorage(new[] {
            top[0] + tsv * TDF / 2 + 20, top[1] - tsv * TD / 2
        }));
        drawingData.DimensionStyles["DatumFeature"] = new DimensioningStorage(new DataStorage(new[] {
            top[0] - tsv * TDF / 2, top[1] - tsv * TD / 2 - 1
        }));
        drawingData.DimensionStyles["ISA"] = new DimensioningStorage(new DataStorage(new[] {
            detail[0] + 3.5, detail[1] + detailLowerLength - 3.75
        }));
        drawingData.DimensionStyles["GA"] = new DimensioningStorage(new DataStorage(new[] {
            detail[0], detail[1] - 2
        }));
        drawingData.DimensionStyles["B"] = new DimensioningStorage(new DataStorage(new[] {
            detail[0], detail[1] - 10
        }));
        drawingData.DimensionStyles["W"] = new DimensioningStorage(new DataStorage(new[] {
            detail[0], detail[1] - 15
        }));
        drawingData.DimensionStyles["GeometricTolerance"] = new DimensioningStorage(new DataStorage(new[] {
            detail[0] - 13.5, detail[1] - 20
        }));
        drawingData.DimensionStyles["GD"] = new DimensioningStorage(new DataStorage(new[] {
            detail[0] - dsv * W / 2 - 10, detail[1] + dsv * GD / 2
        }));
        drawingData.DimensionStyles["GR"] = new DimensioningStorage(new DataStorage(new[] {
            detail[0] + 10, detail[1] + dsv * GD + 5
        }));
        drawingData.DimensionStyles["FA"] = new DimensioningStorage(new DataStorage(new[] {
            side[0] - ssv * TD / 2 - 4, side[1] + 20
        }));
        drawingData.DimensionStyles["BA"] = new DimensioningStorage(new DataStorage(new[] {
            side[0] + ssv * TD / 2 + 4, side[1] + 15
        }));
        drawingData.DimensionStyles["E"] = new DimensioningStorage(new DataStorage(new[] {
            side[0] + ssv * TD / 2 + 2.5, side[1] - 68
        }));
        drawingData.DimensionStyles["FX"] = new DimensioningStorage(new DataStorage(new[] {
            side[0] - ssv * TD / 2 - 10, side[1] - 81.5
        }));
        drawingData.DimensionStyles["F"] = new DimensioningStorage(new DataStorage(new[] {
            section[0], section[1] - 10
        }));
        drawingData.DimensionStyles["FL"] = new DimensioningStorage(new DataStorage(new[] {
            section[0], section[1] - 15
        }));
        drawingData.DimensionStyles["FR"] = new DimensioningStorage(new DataStorage(new[] {
            section[0] - secv * FL / 2 - 10, section[1] + secv * GD / 2
        }));
        drawingData.DimensionStyles["BR"] = new DimensioningStorage(new DataStorage(new[] {
            section[0] + secv * FL / 2 + 10, section[1] + secv * GD / 2
        }));

        return drawingData;
    }

}
