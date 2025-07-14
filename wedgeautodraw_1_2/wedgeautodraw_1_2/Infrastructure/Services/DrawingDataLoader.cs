using System.Globalization;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;
using wedgeautodraw_1_2.Core;

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

        LoadViewScales(drawingData, config);
        LoadViewPositions(drawingData, wedgeData, config, drawingType);
        LoadTablePositions(drawingData, config);
        LoadBreaklineData(drawingData, config);
        LoadTitleBlockInfo(drawingData, wedgeData, config, drawingType);
        LoadExtras(drawingData, config);

        drawingData.DrawingType = drawingType;
        return drawingData;
    }

    private void LoadViewScales(DrawingData data, ConfigLoader config)
    {
        if (config.HasKey(Constants.ConfigKeys.ScalingFSV))
        {
            double fsv = config.GetDouble(Constants.ConfigKeys.ScalingFSV);
            data.ViewScales["Front_view"] = new DataStorage(fsv);
            data.ViewScales["Side_view"] = new DataStorage(fsv);
            data.ViewScales["Top_view"] = new DataStorage(fsv);
        }

        double defaultScale = config.GetDouble(Constants.ConfigKeys.ScalingDSV);
        data.ViewScales["Detail_view"] = new DataStorage(defaultScale);
        data.ViewScales["Section_view"] = new DataStorage(defaultScale);
    }

    private void LoadViewPositions(DrawingData data, WedgeData wedge, ConfigLoader config, DrawingType type)
    {
        if (config.HasKey(Constants.ConfigKeys.FrontViewPosX) && config.HasKey(Constants.ConfigKeys.FrontViewPosY))
        {
            data.ViewPositions["Front_view"] = new DataStorage(new[]
            {
                config.GetDouble(Constants.ConfigKeys.FrontViewPosX),
                config.GetDouble(Constants.ConfigKeys.FrontViewPosY)
            });
        }

        if (config.HasKey(Constants.ConfigKeys.SideViewDX))
        {
            if (type == DrawingType.Production && data.ViewPositions.ContainsKey("Front_view"))
            {
                var frontPos = data.ViewPositions["Front_view"].GetValues(Unit.Millimeter);
                data.ViewPositions["Side_view"] = new DataStorage(new[]
                {
                    frontPos[0] + config.GetDouble(Constants.ConfigKeys.SideViewDX),
                    frontPos[1]
                });
            }
            else
            {
                data.ViewPositions["Side_view"] = new DataStorage(new[]
                {
                    config.GetDouble(Constants.ConfigKeys.SideViewDX),
                    config.GetDouble(Constants.ConfigKeys.SideViewDY)
                });
            }
        }

        if (config.HasKey(Constants.ConfigKeys.TopViewDY))
        {
            if (type == DrawingType.Production && data.ViewPositions.ContainsKey("Side_view"))
            {
                var sidePos = data.ViewPositions["Side_view"].GetValues(Unit.Millimeter);
                data.ViewPositions["Top_view"] = new DataStorage(new[]
                {
                    sidePos[0],
                    sidePos[1] + config.GetDouble(Constants.ConfigKeys.TopViewDY)
                });
            }
            else
            {
                data.ViewPositions["Top_view"] = new DataStorage(new[]
                {
                    config.GetDouble(Constants.ConfigKeys.TopViewDX),
                    config.GetDouble(Constants.ConfigKeys.TopViewDY)
                });
            }
        }

        if (config.HasKey(Constants.ConfigKeys.DetailViewPosX) && config.HasKey(Constants.ConfigKeys.DetailViewPosY))
        {
            data.ViewPositions["Detail_view"] = new DataStorage(new[]
            {
                config.GetDouble(Constants.ConfigKeys.DetailViewPosX),
                config.GetDouble(Constants.ConfigKeys.DetailViewPosY)
            });
        }

        if (config.HasKey(Constants.ConfigKeys.SectionViewPosX) && data.ViewPositions.ContainsKey("Detail_view"))
        {
            double td = wedge.Dimensions.GetOrDefault("TD")?.GetValue(Unit.Millimeter) ?? 0;
            double tdf = wedge.Dimensions.GetOrDefault("TDF")?.GetValue(Unit.Millimeter) ?? 0;
            double fl = wedge.Dimensions.GetOrDefault("FL")?.GetValue(Unit.Millimeter) ?? 0;
            double scale = data.ViewScales["Section_view"].GetValue(Unit.Millimeter);

            double offsetX = config.GetDouble(Constants.ConfigKeys.SectionViewPosX) + scale * (td - tdf) / 2;
            var detailPos = data.ViewPositions["Detail_view"].GetValues(Unit.Millimeter);

            data.ViewPositions["Section_view"] = new DataStorage(new[]
            {
                detailPos[0] + offsetX,
                detailPos[1]
            });
        }
    }

    private void LoadTablePositions(DrawingData data, ConfigLoader config)
    {
        TrySetTable(data, config, "dimension", "dim_table_posX", "dim_table_posY", "dim_table_width");
        TrySetTable(data, config, "how_to_order", "how_to_order_posX", "how_to_order_posY", "how_to_order_width");
        TrySetTable(data, config, "label_as", "label_as_posX", "label_as_posY", "label_as_width");
        TrySetTable(data, config, "polish", "polish_posX", "polish_posY", "polish_width");
        TrySetTable(data, config, "coining_note", "coining_note_posX", "coining_note_posY", "coining_note_width");
    }

    private void LoadBreaklineData(DrawingData data, ConfigLoader config)
    {
        SetBreakline(data, config, "Front_view", "fsv");
        SetBreakline(data, config, "Side_view", "fsv");
        SetBreakline(data, config, "Detail_view", "dsv", false);
        SetBreakline(data, config, "Section_view", "dsv", false);
    }

    private void SetBreakline(DrawingData data, ConfigLoader config, string view, string suffix, bool setUpper = true)
    {
        string lowerKey = $"length_lower_section_{suffix}";
        string upperKey = $"length_upper_section_{suffix}";
        string gapKey = $"breakline_gap_{suffix}";

        if (config.HasKey(lowerKey))
            data.BreaklineData[$"{view}LowerPartLength"] = new DataStorage(config.GetDouble(lowerKey));

        if (setUpper && config.HasKey(upperKey))
            data.BreaklineData[$"{view}UpperPartLength"] = new DataStorage(config.GetDouble(upperKey));
        else if (!setUpper)
            data.BreaklineData[$"{view}UpperPartLength"] = new DataStorage(0);

        if (config.HasKey(gapKey))
            data.BreaklineData[$"{view}BreaklineGap"] = new DataStorage(config.GetDouble(gapKey));
    }

    private void LoadTitleBlockInfo(DrawingData data, WedgeData wedge, ConfigLoader config, DrawingType type)
    {
        data.TitleInfo["number"] = wedge.Metadata["drawing_number"] + "";
        data.Title = wedge.Metadata["wedge_title"] + "";

        data.TitleBlockInfo["Material"] = config.GetString(Constants.ConfigKeys.Material);
        data.TitleBlockInfo["Autor"] = config.GetString(Constants.ConfigKeys.Author);
        data.TitleBlockInfo["DRAWN_BY"] = "AUTODRAW SERVICE";
        data.TitleBlockInfo["COMPANY_NAME"] = "SMALL PRECISION TOOLS";
        data.TitleBlockInfo["TITLE"] = data.Title;
        data.TitleBlockInfo["DRAWING_NUMBER"] = data.TitleInfo["number"] + "-DW";
        data.TitleBlockInfo["ADDRESS"] = "1330 CLEGG STREET PETALUMA, CALIFORNIA 94954";
        data.TitleBlockInfo["TYPE"] = type.ToString().ToUpperInvariant();
        data.TitleBlockInfo["SCALING_FRONT_SIDE_TOP_VIEW"] = config.HasKey(Constants.ConfigKeys.ScalingFSV) ? config.GetDouble(Constants.ConfigKeys.ScalingFSV).ToString() : "";
        data.TitleBlockInfo["SCALING_DETAIL_SECTION_VIEW"] = config.GetDouble(Constants.ConfigKeys.ScalingDSV).ToString();
        data.TitleBlockInfo["DRAWN_ON"] = DateTime.Now.ToString("MM-dd-yy");

        data.HowToOrderInfo["number"] = data.TitleInfo["number"];
        data.HowToOrderInfo["packaging"] = config.GetString(Constants.ConfigKeys.Packaging);
    }

    private void LoadExtras(DrawingData data, ConfigLoader config)
    {
        string engrave = config.GetString(Constants.ConfigKeys.Engrave);
        if (!string.IsNullOrWhiteSpace(engrave))
            data.LabelAsItems = engrave.Split('¶');

        string polish = config.GetString(Constants.ConfigKeys.PolishText);
        if (!string.IsNullOrWhiteSpace(polish))
            data.PolishItems = polish.Split('¶');

        string dimKeys = config.GetString(Constants.ConfigKeys.DimensionKeysInTable);
        if (!string.IsNullOrWhiteSpace(dimKeys))
            data.DimensionKeysInTable = dimKeys.Split(',');
    }

    private void TrySetTable(DrawingData data, ConfigLoader config, string name, string xKey, string yKey, string wKey)
    {
        if (config.HasKey(xKey) && config.HasKey(yKey) && config.HasKey(wKey))
        {
            data.TablePositions[name] = new DataStorage(new[]
            {
                config.GetDouble(xKey),
                config.GetDouble(yKey),
                config.GetDouble(wKey)
            });
        }
    }
}
