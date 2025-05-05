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
                return new[] { front[0] + fsv * TD / 2 + 4, front[1] * 45};
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
                var breakline = drawing.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Millimeter);
                var y = detail[1] - (breakline + lowerPartLength) / 2;
                return new[] { detail[0] + 3.5, y + lowerPartLength - 3.75 };
            }
        },
        ["GA"] = new DimensionRule
        {
            BasedOnView = Constants.DetailView,
            CalculatePosition = (wedge, drawing) =>
            {
                var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                var breakline = drawing.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Millimeter);
                var y = detail[1] - (breakline + lowerPartLength) / 2;
                return new[] { detail[0], y - 2 };
            }
        },
        ["B"] = new DimensionRule
        {
            BasedOnView = Constants.DetailView,
            CalculatePosition = (wedge, drawing) =>
            {
                var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                var breakline = drawing.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Millimeter);
                var y = detail[1] - (breakline + lowerPartLength) / 2;
                return new[] { detail[0], y - 10 };
            }
        },
        ["W"] = new DimensionRule
        {
            BasedOnView = Constants.DetailView,
            CalculatePosition = (wedge, drawing) =>
            {
                var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                var breakline = drawing.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Millimeter);
                var y = detail[1] - (breakline + lowerPartLength) / 2;
                return new[] { detail[0], y - 15 };
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
                var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                var breakline = drawing.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Millimeter);
                var y = detail[1] - (breakline + lowerPartLength) / 2;
                return new[] { detail[0] - dsv * W / 2 - 10, y + dsv * GD / 2 };
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
                var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                var breakline = drawing.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Millimeter);
                var y = detail[1] - (breakline + lowerPartLength) / 2;
                return new[] { detail[0] + 10, y + dsv * GD + 5 };
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
            BasedOnView = "Section_view",
            CalculatePosition = (wedge, drawing) =>
            {
                var section = drawing.ViewPositions["Section_view"].GetValues(Unit.Millimeter);
                var scale = drawing.ViewScales["Section_view"].GetValue(Unit.Millimeter);
                double FL = wedge.Dimensions["FL"].GetValue(Unit.Millimeter);
                double TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                double TDF = wedge.Dimensions["TDF"].GetValue(Unit.Millimeter);
                var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                var breakline = drawing.BreaklineData["Section_viewBreaklineGap"].GetValue(Unit.Millimeter);
                double secv_y = section[1] - (breakline + lowerPartLength) / 2;

                double secv_x = section[0]; // fallback
                if (wedge.Dimensions.TryGet("FX", out var fxStorage) && fxStorage != null)
                {
                    double FX = fxStorage.GetValue(Unit.Millimeter);
                    if (FX != 0 && !double.IsNaN(FX))
                        secv_x = section[0] - scale * (TDF / 2 - FX - FL / 2);
                    else
                        secv_x = section[0] - scale * (TD - TDF) / 2;
                }
                else
                {
                    secv_x = section[0] - scale * (TD - TDF) / 2;
                }

                return new[] { secv_x, secv_y -10 };
            }
        },
        ["FL"] = new DimensionRule
        {
            BasedOnView = Constants.SectionView,
            CalculatePosition = (wedge, drawing) =>
            {
                var section = drawing.ViewPositions["Section_view"].GetValues(Unit.Millimeter);
                var scale = drawing.ViewScales["Section_view"].GetValue(Unit.Millimeter);
                double FL = wedge.Dimensions["FL"].GetValue(Unit.Millimeter);
                double TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                double TDF = wedge.Dimensions["TDF"].GetValue(Unit.Millimeter);
                var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                var breakline = drawing.BreaklineData["Section_viewBreaklineGap"].GetValue(Unit.Millimeter);
                double secv_y = section[1] - (breakline + lowerPartLength) / 2;

                double secv_x = section[0];
                if (wedge.Dimensions.TryGet("FX", out var fxStorage) && fxStorage != null)
                {
                    double FX = fxStorage.GetValue(Unit.Millimeter);
                    if (FX != 0 && !double.IsNaN(FX))
                        secv_x = section[0] - scale * (TDF / 2 - FX - FL / 2);
                    else
                        secv_x = section[0] - scale * (TD - TDF) / 2;
                }
                else
                {
                    secv_x = section[0] - scale * (TD - TDF) / 2;
                }

                return new[] { secv_x, secv_y - 15 };
            }
        },
        ["FR"] = new DimensionRule
        {
            BasedOnView = Constants.SectionView,
            CalculatePosition = (wedge, drawing) =>
            {
                var section = drawing.ViewPositions["Section_view"].GetValues(Unit.Millimeter);
                var scale = drawing.ViewScales["Section_view"].GetValue(Unit.Millimeter);
                double FL = wedge.Dimensions["FL"].GetValue(Unit.Millimeter);
                double GD = wedge.Dimensions["GD"].GetValue(Unit.Millimeter);
                double TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                double TDF = wedge.Dimensions["TDF"].GetValue(Unit.Millimeter);
                var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                var breakline = drawing.BreaklineData["Section_viewBreaklineGap"].GetValue(Unit.Millimeter);
                double secv_y = section[1] - (breakline + lowerPartLength) / 2;

                double secv_x = section[0];
                if (wedge.Dimensions.TryGet("FX", out var fxStorage) && fxStorage != null)
                {
                    double FX = fxStorage.GetValue(Unit.Millimeter);
                    if (FX != 0 && !double.IsNaN(FX))
                        secv_x = section[0] - scale * (TDF / 2 - FX - FL / 2);
                    else
                        secv_x = section[0] - scale * (TD - TDF) / 2;
                }
                else
                {
                    secv_x = section[0] - scale * (TD - TDF) / 2;
                }

                return new[] { secv_x - scale * FL / 2 - 10, secv_y + scale * GD / 2 };
            }
        },
        ["BR"] = new DimensionRule
        {
            BasedOnView = Constants.SectionView,
            CalculatePosition = (wedge, drawing) =>
            {
                var section = drawing.ViewPositions["Section_view"].GetValues(Unit.Millimeter);
                var scale = drawing.ViewScales["Section_view"].GetValue(Unit.Millimeter);
                double FL = wedge.Dimensions["FL"].GetValue(Unit.Millimeter);
                double GD = wedge.Dimensions["GD"].GetValue(Unit.Millimeter);
                double TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                double TDF = wedge.Dimensions["TDF"].GetValue(Unit.Millimeter);
                var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                var breakline = drawing.BreaklineData["Section_viewBreaklineGap"].GetValue(Unit.Millimeter);
                double secv_y = section[1] - (breakline + lowerPartLength) / 2;

                double secv_x = section[0];
                if (wedge.Dimensions.TryGet("FX", out var fxStorage) && fxStorage != null)
                {
                    double FX = fxStorage.GetValue(Unit.Millimeter);
                    if (FX != 0 && !double.IsNaN(FX))
                        secv_x = section[0] - scale * (TDF / 2 - FX - FL / 2);
                    else
                        secv_x = section[0] - scale * (TD - TDF) / 2;
                }
                else
                {
                    secv_x = section[0] - scale * (TD - TDF) / 2;
                }

                return new[] { secv_x + scale * FL / 2 + 10, secv_y + scale * GD / 2 };
            }
        }

    };
}
