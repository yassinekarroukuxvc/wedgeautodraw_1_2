using System.Globalization;
using System.Text.RegularExpressions;
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

        double defaultScale = Get("scaling_dsv");
        double W_value = wedgeData.Dimensions.ContainsKey("W") ? wedgeData.Dimensions["W"].GetValue(Unit.Millimeter) : 10.0;
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

        drawingData.ViewPositions["Front_view"] = new DataStorage(new[] { Get("front_view_posX"), Get("front_view_posY") });
        drawingData.ViewPositions["Side_view"] = new DataStorage(new[] {
            Get("front_view_posX") + Get("side_view_dX"), Get("front_view_posY") });
        drawingData.ViewPositions["Top_view"] = new DataStorage(new[] {
            Get("front_view_posX") + Get("side_view_dX"), Get("front_view_posY") + Get("top_view_dY") });
        drawingData.ViewPositions["Detail_view"] = new DataStorage(new[] { Get("detail_view_posX"), Get("detail_view_posY") });
        drawingData.ViewPositions["Section_view"] = new DataStorage(new[] {
            Get("detail_view_posX") + sectionViewCenterX, Get("detail_view_posY") });

        drawingData.TablePositions["dimension"] = new DataStorage(new[] {
            Get("dim_table_posX"), Get("dim_table_posY"), Get("dim_table_width") });
        drawingData.TablePositions["how_to_order"] = new DataStorage(new[] {
            Get("how_to_order_posX"), Get("how_to_order_posY"), Get("how_to_order_width") });
        drawingData.TablePositions["label_as"] = new DataStorage(new[] {
            Get("label_as_posX"), Get("label_as_posY"), Get("label_as_width") });
        drawingData.TablePositions["polish"] = new DataStorage(new[] {
            Get("polish_posX"), Get("polish_posY"), Get("polish_width") });

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

        drawingData.TitleInfo["info"] = GetStr("how_to_order_info");
        drawingData.TitleInfo["number"] = (string)wedgeData.Metadata["drawing_number"];
        drawingData.Title = wedgeData.Metadata["wedge_title"] + "";

        drawingData.TitleBlockInfo["Material"] = GetStr("material");
        drawingData.TitleBlockInfo["Autor"] = GetStr("author");
        drawingData.TitleBlockInfo["DRAWN_BY"] = "AUTODRAW SERVICE";
        drawingData.TitleBlockInfo["COMPANY_NAME"] = "SMALL PRECISION TOOLS";
        drawingData.TitleBlockInfo["TITLE"] = drawingData.Title;
        drawingData.TitleBlockInfo["DRAWING_NUMBER"] = drawingData.TitleInfo["number"] + "-DW";
        drawingData.TitleBlockInfo["ADDRESS"] = "1330 CLEGG STREET PETALUMA, CALIFORNIA 94954";
        drawingData.TitleBlockInfo["TYPE"] = drawingType.ToString().ToUpperInvariant();
        drawingData.TitleBlockInfo["SCALING_FRONT_SIDE_TOP_VIEW"] = Get("scaling_fsv").ToString();
        drawingData.TitleBlockInfo["SCALING_DETAIL_SECTION_VIEW"] = Get("scaling_dsv").ToString();
        drawingData.TitleBlockInfo["DRAWN_ON"] = DateTime.Now.ToString("MM-dd-yy");

        drawingData.HowToOrderInfo["number"] = drawingData.TitleInfo["number"];
        drawingData.HowToOrderInfo["packaging"] = GetStr("packaging");

        drawingData.DrawingType = drawingType;
        drawingData.LabelAsItems = config.GetStringArray("engrave");
        drawingData.PolishItems = config.GetStringArray("polish_text");
        drawingData.DimensionKeysInTable = config.GetString("dimension_keys_in_table").Split(',');

        DynamicDimensionStyler.ApplyDynamicStyles(drawingData, wedgeData);
        return drawingData;
    }
}

