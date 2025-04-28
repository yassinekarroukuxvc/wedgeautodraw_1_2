using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Infrastructure.Services;

public static class DrawingDataStyler
{
    public static void ApplyDimensionStyles(DrawingData drawingData, WedgeData wedgeData)
    {
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
        var F = wedgeData.Dimensions["F"].GetValue(Unit.Millimeter);
        var TL = wedgeData.Dimensions["TL"].GetValue(Unit.Millimeter);

        var detailLowerLength = drawingData.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);

        drawingData.DimensionStyles["TL"] = new DimensioningStorage(new DataStorage(new[] {
            front[0] - fsv * TD / 2 - 7.5, front[1]
        }));

        drawingData.DimensionStyles["EngravingStart"] = new DimensioningStorage(new DataStorage(new[] {
            front[0] + fsv * TD / 2 + 4, (TL/2) * 1000
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
            detail[0] - 13.5, detail[1] - 70
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
            section[0], section[1] - 55
        }));

        drawingData.DimensionStyles["FL"] = new DimensioningStorage(new DataStorage(new[] {
            section[0], section[1] - 65
        }));

        drawingData.DimensionStyles["FR"] = new DimensioningStorage(new DataStorage(new[] {
            section[0] - secv * FL / 2, section[1] + secv * GD / 3
        }));

        drawingData.DimensionStyles["BR"] = new DimensioningStorage(new DataStorage(new[] {
            section[0] + secv * FL / 2, section[1] + secv * GD / 3
        }));
    }
}
