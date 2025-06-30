using System.Globalization;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services;

public class DrawingDataLoader : IDrawingDataLoader
{
    private readonly string _equationFilePath;

    public DrawingDataLoader(string equationFilePath)
    {
        _equationFilePath = equationFilePath;
    }

    public DrawingData LoadDrawingData(WedgeData wedgeData, string configFilePath, DrawingType drawingType)
    {
        var drawingData = new DrawingData();
        var config = new ConfigLoader(configFilePath, drawingType);

        double Get(string key) => config.GetDouble(key);
        string GetStr(string key) => config.GetString(key);
        string[] GetArray(string key) => config.GetStringArray(key);
        bool Has(string key) => !double.IsNaN(Get(key));

        // === View Scales ===
        if (Has("scaling_fsv"))
        {
            double fsv = Get("scaling_fsv");
            drawingData.ViewScales["Front_view"] = new DataStorage(fsv);
            drawingData.ViewScales["Side_view"] = new DataStorage(fsv);
            drawingData.ViewScales["Top_view"] = new DataStorage(fsv);
        }

        double defaultScale = Get("scaling_dsv");
        /*double W_value = wedgeData.Dimensions.ContainsKey("W") ? wedgeData.Dimensions["W"].GetValue(Unit.Millimeter) : 10.0;
        double adjustedScale = W_value >= 0.7 ? Math.Max(defaultScale * (1.0 / W_value), 0.2) : defaultScale;
        adjustedScale = Math.Round(adjustedScale, 3);*/
        drawingData.ViewScales["Detail_view"] = new DataStorage(defaultScale);
        drawingData.ViewScales["Section_view"] = new DataStorage(defaultScale);

        // === View Positions ===
        if (Has("front_view_posX") && Has("front_view_posY"))
        {
            drawingData.ViewPositions["Front_view"] = new DataStorage(new[] { Get("front_view_posX"), Get("front_view_posY") });
        }

        if (Has("side_view_dX") && drawingData.ViewPositions.ContainsKey("Front_view"))
        {
            
            if(drawingType == DrawingType.Production)
            {
                var basePos = drawingData.ViewPositions["Front_view"].GetValues(Unit.Millimeter);
                drawingData.ViewPositions["Side_view"] = new DataStorage(new[] { basePos[0] + Get("side_view_dX"), basePos[1] });
            }
            else
            {
                drawingData.ViewPositions["Side_view"] = new DataStorage(new[] {Get("side_view_dX"), Get("side_view_dY") });
            }
        }

        if (Has("top_view_dY") && drawingData.ViewPositions.ContainsKey("Side_view"))
        {
            
            if (drawingType == DrawingType.Production)
            {
                var basePos = drawingData.ViewPositions["Side_view"].GetValues(Unit.Millimeter);
                drawingData.ViewPositions["Top_view"] = new DataStorage(new[] { basePos[0], basePos[1] + Get("top_view_dY") });
            }
            else
            {
                drawingData.ViewPositions["Top_view"] = new DataStorage(new[] { Get("top_view_dX") , Get("top_view_dY") });
            }
        }

        if (Has("detail_view_posX") && Has("detail_view_posY"))
        {
            drawingData.ViewPositions["Detail_view"] = new DataStorage(new[] { Get("detail_view_posX"), Get("detail_view_posY") });
        }

        if (Has("section_view_posX") && drawingData.ViewPositions.ContainsKey("Detail_view"))
        {
            double fx = 0.0;
            double td = wedgeData.Dimensions.ContainsKey("TD") ? wedgeData.Dimensions["TD"].GetValue(Unit.Millimeter) : 0;
            double tdf = wedgeData.Dimensions.ContainsKey("TDF") ? wedgeData.Dimensions["TDF"].GetValue(Unit.Millimeter) : 0;
            double fl = wedgeData.Dimensions.ContainsKey("FL") ? wedgeData.Dimensions["FL"].GetValue(Unit.Millimeter) : 0;
            double scale = drawingData.ViewScales["Section_view"].GetValue(Unit.Millimeter);

            double offsetX = fx == 0.0 || double.IsNaN(fx)
                ? Get("section_view_posX") + scale * (td - tdf) / 2
                : Get("section_view_posX") + scale * ((tdf - fl) / 2 - fx);

            var detailPos = drawingData.ViewPositions["Detail_view"].GetValues(Unit.Millimeter);
            drawingData.ViewPositions["Section_view"] = new DataStorage(new[] { detailPos[0] + offsetX, detailPos[1] });
        }

        // === Table Positions ===
        TrySetTable("dimension", "dim_table_posX", "dim_table_posY", "dim_table_width");
        TrySetTable("how_to_order", "how_to_order_posX", "how_to_order_posY", "how_to_order_width");
        TrySetTable("label_as", "label_as_posX", "label_as_posY", "label_as_width");
        TrySetTable("polish", "polish_posX", "polish_posY", "polish_width");
        TrySetTable("coining_note", "coining_note_posX", "coining_note_posY", "coining_note_width");

        // === Breakline Data ===
        void SetBreakline(string view, string suffix, bool setUpper = true)
        {
            if (Has($"length_lower_section_{suffix}"))
                drawingData.BreaklineData[$"{view}LowerPartLength"] = new DataStorage(Get($"length_lower_section_{suffix}"));

            if (setUpper && Has($"length_upper_section_{suffix}"))
                drawingData.BreaklineData[$"{view}UpperPartLength"] = new DataStorage(Get($"length_upper_section_{suffix}"));
            else if (!setUpper)
                drawingData.BreaklineData[$"{view}UpperPartLength"] = new DataStorage(0);

            if (Has($"breakline_gap_{suffix}"))
                drawingData.BreaklineData[$"{view}BreaklineGap"] = new DataStorage(Get($"breakline_gap_{suffix}"));
        }

        SetBreakline("Front_view", "fsv");
        SetBreakline("Side_view", "fsv");
        SetBreakline("Detail_view", "dsv", false);
        SetBreakline("Section_view", "dsv", false);

        // === Title Block Info ===
        drawingData.TitleInfo["number"] = wedgeData.Metadata["drawing_number"] + "";
        drawingData.Title = wedgeData.Metadata["wedge_title"] + "";

        drawingData.TitleBlockInfo["Material"] = GetStr("material");
        drawingData.TitleBlockInfo["Autor"] = GetStr("author");
        drawingData.TitleBlockInfo["DRAWN_BY"] = "AUTODRAW SERVICE";
        drawingData.TitleBlockInfo["COMPANY_NAME"] = "SMALL PRECISION TOOLS";
        drawingData.TitleBlockInfo["TITLE"] = drawingData.Title;
        drawingData.TitleBlockInfo["DRAWING_NUMBER"] = drawingData.TitleInfo["number"] + "-DW";
        drawingData.TitleBlockInfo["ADDRESS"] = "1330 CLEGG STREET PETALUMA, CALIFORNIA 94954";
        drawingData.TitleBlockInfo["TYPE"] = drawingType.ToString().ToUpperInvariant();
        drawingData.TitleBlockInfo["SCALING_FRONT_SIDE_TOP_VIEW"] = Has("scaling_fsv") ? Get("scaling_fsv").ToString() : "";
        drawingData.TitleBlockInfo["SCALING_DETAIL_SECTION_VIEW"] = Get("scaling_dsv").ToString();
        drawingData.TitleBlockInfo["DRAWN_ON"] = DateTime.Now.ToString("MM-dd-yy");

        drawingData.HowToOrderInfo["number"] = drawingData.TitleInfo["number"];
        drawingData.HowToOrderInfo["packaging"] = GetStr("packaging");

        // === Extras ===
        string engrave = GetStr("engrave");
        if (!string.IsNullOrWhiteSpace(engrave))
            drawingData.LabelAsItems = engrave.Split('¶');

        string polish = GetStr("polish_text");
        if (!string.IsNullOrWhiteSpace(polish))
            drawingData.PolishItems = polish.Split('¶');

        string dimKeys = GetStr("dimension_keys_in_table");
        if (!string.IsNullOrWhiteSpace(dimKeys))
            drawingData.DimensionKeysInTable = dimKeys.Split(',');

        drawingData.DrawingType = drawingType;
        //DynamicDimensionStyler.ApplyDynamicStyles(drawingData, wedgeData);
        return drawingData;

        // === Helper for Tables ===
        void TrySetTable(string name, string xKey, string yKey, string wKey)
        {
            if (Has(xKey) && Has(yKey) && Has(wKey))
                drawingData.TablePositions[name] = new DataStorage(new[] { Get(xKey), Get(yKey), Get(wKey) });
        }
    }
}
