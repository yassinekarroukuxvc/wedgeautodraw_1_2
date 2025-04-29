using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Core.Enums;
namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public static class DimensionRules
{
    public static readonly Dictionary<string, DimensionRule> Rules = new()
    {
        ["TL"] = new DimensionRule
        {
            BasedOnView = Constants.FrontView,
            CalculatePosition = (wedge, drawing) =>
            {
                var front = drawing.ViewPositions[Constants.FrontView].GetValues(Unit.Millimeter);
                var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                var fsv = drawing.ViewScales[Constants.FrontView].GetValue(Unit.Millimeter);
                return new[] { front[0] - fsv * TD / 2 - 7.5, front[1] };
            }
        },
        ["EngravingStart"] = new DimensionRule
        {
            BasedOnView = Constants.FrontView,
            CalculatePosition = (wedge, drawing) =>
            {
                var front = drawing.ViewPositions[Constants.FrontView].GetValues(Unit.Millimeter);
                var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                var TL = wedge.Dimensions["TL"].GetValue(Unit.Millimeter);
                var fsv = drawing.ViewScales[Constants.FrontView].GetValue(Unit.Millimeter);
                return new[] { front[0] + fsv * TD / 2 + 4, (TL / 2) * 1000 };
            }
        },
        ["TDF"] = new DimensionRule
        {
            BasedOnView = Constants.TopView,
            CalculatePosition = (wedge, drawing) =>
            {
                var top = drawing.ViewPositions[Constants.TopView].GetValues(Unit.Millimeter);
                var TDF = wedge.Dimensions["TDF"].GetValue(Unit.Millimeter);
                var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                var tsv = drawing.ViewScales[Constants.TopView].GetValue(Unit.Millimeter);
                return new[] { top[0] + tsv * TDF / 2 + 20, top[1] + tsv * TD / 2 + 3 };
            }
        },
        ["TD"] = new DimensionRule
        {
            BasedOnView = Constants.TopView,
            CalculatePosition = (wedge, drawing) =>
            {
                var top = drawing.ViewPositions[Constants.TopView].GetValues(Unit.Millimeter);
                var TDF = wedge.Dimensions["TDF"].GetValue(Unit.Millimeter);
                var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                var tsv = drawing.ViewScales[Constants.TopView].GetValue(Unit.Millimeter);
                return new[] { top[0] + tsv * TDF / 2 + 20, top[1] - tsv * TD / 2 };
            }
        },
        ["DatumFeature"] = new DimensionRule
        {
            BasedOnView = Constants.TopView,
            CalculatePosition = (wedge, drawing) =>
            {
                var top = drawing.ViewPositions[Constants.TopView].GetValues(Unit.Millimeter);
                var TDF = wedge.Dimensions["TDF"].GetValue(Unit.Millimeter);
                var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                var tsv = drawing.ViewScales[Constants.TopView].GetValue(Unit.Millimeter);
                return new[] { top[0] - tsv * TDF / 2, top[1] - tsv * TD / 2 - 1 };
            }
        },
        ["ISA"] = new DimensionRule
        {
            BasedOnView = Constants.DetailView,
            CalculatePosition = (wedge, drawing) =>
            {
                var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                return new[] { detail[0] + 3.5, detail[1] + lowerPartLength - 3.75 };
            }
        },
        ["GA"] = new DimensionRule
        {
            BasedOnView = Constants.DetailView,
            CalculatePosition = (wedge, drawing) =>
            {
                var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                return new[] { detail[0], detail[1] - 2 };
            }
        },
        ["B"] = new DimensionRule
        {
            BasedOnView = Constants.DetailView,
            CalculatePosition = (wedge, drawing) =>
            {
                var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                return new[] { detail[0], detail[1] - 10 };
            }
        },
        ["W"] = new DimensionRule
        {
            BasedOnView = Constants.DetailView,
            CalculatePosition = (wedge, drawing) =>
            {
                var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                return new[] { detail[0], detail[1] - 15 };
            }
        },
        ["GeometricTolerance"] = new DimensionRule
        {
            BasedOnView = Constants.DetailView,
            CalculatePosition = (wedge, drawing) =>
            {
                var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                return new[] { detail[0] - 13.5, detail[1] - 70 };
            }
        },
        ["GD"] = new DimensionRule
        {
            BasedOnView = Constants.DetailView,
            CalculatePosition = (wedge, drawing) =>
            {
                var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                var W = wedge.Dimensions["W"].GetValue(Unit.Millimeter);
                var GD = wedge.Dimensions["GD"].GetValue(Unit.Millimeter);
                var dsv = drawing.ViewScales[Constants.DetailView].GetValue(Unit.Millimeter);
                return new[] { detail[0] - dsv * W / 2 - 10, detail[1] + dsv * GD / 2 };
            }
        },
        ["GR"] = new DimensionRule
        {
            BasedOnView = Constants.DetailView,
            CalculatePosition = (wedge, drawing) =>
            {
                var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                var GD = wedge.Dimensions["GD"].GetValue(Unit.Millimeter);
                var dsv = drawing.ViewScales[Constants.DetailView].GetValue(Unit.Millimeter);
                return new[] { detail[0] + 10, detail[1] + dsv * GD + 5 };
            }
        },
        ["FA"] = new DimensionRule
        {
            BasedOnView = Constants.SideView,
            CalculatePosition = (wedge, drawing) =>
            {
                var side = drawing.ViewPositions[Constants.SideView].GetValues(Unit.Millimeter);
                var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                var ssv = drawing.ViewScales[Constants.SideView].GetValue(Unit.Millimeter);
                return new[] { side[0] - ssv * TD / 2 - 4, side[1] + 20 };
            }
        },
        ["BA"] = new DimensionRule
        {
            BasedOnView = Constants.SideView,
            CalculatePosition = (wedge, drawing) =>
            {
                var side = drawing.ViewPositions[Constants.SideView].GetValues(Unit.Millimeter);
                var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                var ssv = drawing.ViewScales[Constants.SideView].GetValue(Unit.Millimeter);
                return new[] { side[0] + ssv * TD / 2 + 4, side[1] + 15 };
            }
        },
        ["E"] = new DimensionRule
        {
            BasedOnView = Constants.SideView,
            CalculatePosition = (wedge, drawing) =>
            {
                var side = drawing.ViewPositions[Constants.SideView].GetValues(Unit.Millimeter);
                var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                var ssv = drawing.ViewScales[Constants.SideView].GetValue(Unit.Millimeter);
                return new[] { side[0] + ssv * TD / 2 + 2.5, side[1] - 68 };
            }
        },
        ["FX"] = new DimensionRule
        {
            BasedOnView = Constants.SideView,
            CalculatePosition = (wedge, drawing) =>
            {
                var side = drawing.ViewPositions[Constants.SideView].GetValues(Unit.Millimeter);
                var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                var ssv = drawing.ViewScales[Constants.SideView].GetValue(Unit.Millimeter);
                return new[] { side[0] - ssv * TD / 2 - 10, side[1] - 81.5 };
            }
        },
        ["F"] = new DimensionRule
        {
            BasedOnView = Constants.SectionView,
            CalculatePosition = (wedge, drawing) =>
            {
                var section = drawing.ViewPositions[Constants.SectionView].GetValues(Unit.Millimeter);
                return new[] { section[0], section[1] - 55 };
            }
        },
        ["FL"] = new DimensionRule
        {
            BasedOnView = Constants.SectionView,
            CalculatePosition = (wedge, drawing) =>
            {
                var section = drawing.ViewPositions[Constants.SectionView].GetValues(Unit.Millimeter);
                return new[] { section[0], section[1] - 65 };
            }
        },
        ["FR"] = new DimensionRule
        {
            BasedOnView = Constants.SectionView,
            CalculatePosition = (wedge, drawing) =>
            {
                var section = drawing.ViewPositions[Constants.SectionView].GetValues(Unit.Millimeter);
                var FL = wedge.Dimensions["FL"].GetValue(Unit.Millimeter);
                var GD = wedge.Dimensions["GD"].GetValue(Unit.Millimeter);
                var secv = drawing.ViewScales[Constants.SectionView].GetValue(Unit.Millimeter);
                return new[] { section[0] - secv * FL / 2, section[1] + secv * GD / 3 };
            }
        },
        ["BR"] = new DimensionRule
        {
            BasedOnView = Constants.SectionView,
            CalculatePosition = (wedge, drawing) =>
            {
                var section = drawing.ViewPositions[Constants.SectionView].GetValues(Unit.Millimeter);
                var FL = wedge.Dimensions["FL"].GetValue(Unit.Millimeter);
                var GD = wedge.Dimensions["GD"].GetValue(Unit.Millimeter);
                var secv = drawing.ViewScales[Constants.SectionView].GetValue(Unit.Millimeter);
                return new[] { section[0] + secv * FL / 2, section[1] + secv * GD / 3 };
            }
        }
    };
}
