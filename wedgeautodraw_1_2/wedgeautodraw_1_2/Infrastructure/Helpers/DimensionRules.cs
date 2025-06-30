using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public static class DimensionRules
{
    public static Dictionary<string, DimensionLayoutRule> GetRules(DrawingType drawingType)
    {
        return drawingType switch
        {
            DrawingType.Production => GetProductionRules(),
            DrawingType.Overlay => GetOverlayRules(),
            _ => new Dictionary<string, DimensionLayoutRule>()
        };
    }

    private static Dictionary<string, DimensionLayoutRule> GetProductionRules()
    {
        return new Dictionary<string, DimensionLayoutRule>
        {
            ["TL"] = new DimensionLayoutRule
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
            ["EngravingStart"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.FrontView,
                CalculatePosition = (wedge, drawing) =>
                {
                    var front = drawing.ViewPositions[Constants.FrontView].GetValues(Unit.Millimeter);
                    var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                    var TL = wedge.Dimensions["TL"].GetValue(Unit.Millimeter);
                    var fsv = drawing.ViewScales[Constants.FrontView].GetValue(Unit.Millimeter);
                    return new[] { front[0] + fsv * TD / 2 + 10, front[1] + 40 };
                }
            },
            ["D2"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.FrontView,
                CalculatePosition = (wedge, drawing) =>
                {
                    var front = drawing.ViewPositions[Constants.FrontView].GetValues(Unit.Millimeter);
                    var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                    var fsv = drawing.ViewScales[Constants.FrontView].GetValue(Unit.Millimeter);
                    return new[] { front[0] - fsv * TD / 2 + 25, front[1] - 25 };
                }
            },
            ["VW"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.FrontView,
                CalculatePosition = (wedge, drawing) =>
                {
                    var front = drawing.ViewPositions[Constants.FrontView].GetValues(Unit.Millimeter);
                    var lowerPartLength = drawing.BreaklineData["Front_viewLowerPartLength"].GetValue(Unit.Millimeter);
                    var upperPartLength = drawing.BreaklineData["Front_viewUpperPartLength"].GetValue(Unit.Millimeter);
                    var breakline = drawing.BreaklineData["Front_viewBreaklineGap"].GetValue(Unit.Millimeter);
                    var half = (lowerPartLength + breakline + upperPartLength)/2;
                    
                    var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                    var fsv = drawing.ViewScales[Constants.FrontView].GetValue(Unit.Millimeter);
                    var TL = wedge.Dimensions["TL"].GetValue(Unit.Millimeter);
                    return new[] { front[0] - fsv * TD / 2 - 8, front[1] - half - 8 };
                }
            },
            ["TDF"] = new DimensionLayoutRule
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
            ["TD"] = new DimensionLayoutRule
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
            ["DatumFeature"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.TopView,
                CalculatePosition = (wedge, drawing) =>
                {
                    var top = drawing.ViewPositions[Constants.TopView].GetValues(Unit.Millimeter);
                    var TDF = wedge.Dimensions["TDF"].GetValue(Unit.Millimeter);
                    var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                    var tsv = drawing.ViewScales[Constants.TopView].GetValue(Unit.Millimeter);
                    return new[] { top[0] - tsv * TDF / 2 - 20, top[1] - tsv * TD / 2 - 1 };
                }
            },
            ["ISA"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.DetailView,
                CalculatePosition = (wedge, drawing) =>
                {
                    var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                    var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                    var breakline = drawing.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Millimeter);
                    var y = detail[1] - (breakline + lowerPartLength) / 2;
                    return new[] { detail[0] + 3.5, y + lowerPartLength -2 };
                }
            },
            ["GA"] = new DimensionLayoutRule
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
            ["B"] = new DimensionLayoutRule
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
            ["W"] = new DimensionLayoutRule
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
            ["GeometricTolerance"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.DetailView,
                CalculatePosition = (wedge, drawing) =>
                {
                    var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                    return new[] { detail[0] - 13.5, detail[1] - 70 };
                }
            },
            ["GD"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.DetailView,
                CalculatePosition = (wedge, drawing) =>
                {
                    var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                    var W = wedge.Dimensions["W"].GetValue(Unit.Millimeter);
                    var GD = wedge.Dimensions["GD"].GetValue(Unit.Millimeter);
                    var td = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                    var dsv = drawing.ViewScales[Constants.DetailView].GetValue(Unit.Millimeter);
                    var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                    var breakline = drawing.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Millimeter);
                    var y = detail[1] - (breakline + lowerPartLength) / 2;
                    return new[] { detail[0] - dsv * (td / 2) - 15, y + dsv * GD / 2 };
                }
            },
            ["D1"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.DetailView,
                CalculatePosition = (wedge, drawing) =>
                {
                    var detail = drawing.ViewPositions[Constants.DetailView].GetValues(Unit.Millimeter);
                    var W = wedge.Dimensions["W"].GetValue(Unit.Millimeter);
                    var GD = wedge.Dimensions["GD"].GetValue(Unit.Millimeter);
                    var td = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                    var dsv = drawing.ViewScales[Constants.DetailView].GetValue(Unit.Millimeter);
                    var lowerPartLength = drawing.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Millimeter);
                    var breakline = drawing.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Millimeter);
                    var y = detail[1] - (breakline + lowerPartLength) / 2;
                    return new[] { detail[0] - (W/2 *dsv) - 20, y + dsv * GD / 2 };
                }
            },
            ["GR"] = new DimensionLayoutRule
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
                    return new[] { detail[0] - 5, y + dsv * GD + 20 };
                }
            },
            ["FA"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.SideView,
                CalculatePosition = (wedge, drawing) =>
                {
                    var side = drawing.ViewPositions[Constants.SideView].GetValues(Unit.Millimeter);
                    var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                    var ssv = drawing.ViewScales[Constants.SideView].GetValue(Unit.Millimeter);
                    double FA = wedge.Dimensions.TryGet("FA", out var faVal) ? faVal.GetValue(Unit.Degree) : double.NaN;

                    if (!double.IsNaN(FA) && FA < 6)
                        return new[] { side[0] - ssv * TD / 2 - 4, side[1] + 70 };
                    else
                        return new[] { side[0] - ssv * TD / 2 - 4, side[1] + 20 };
                }
            },
            ["BA"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.SideView,
                CalculatePosition = (wedge, drawing) =>
                {
                    var side = drawing.ViewPositions[Constants.SideView].GetValues(Unit.Millimeter);
                    var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                    var ssv = drawing.ViewScales[Constants.SideView].GetValue(Unit.Millimeter);
                    double BA = wedge.Dimensions.TryGet("BA", out var baVal) ? baVal.GetValue(Unit.Degree) : double.NaN;

                    if (!double.IsNaN(BA) && BA < 6)
                        return new[] { side[0] - ssv * TD / 2 + 4, side[1] + 55 };
                    else
                        return new[] { side[0] + ssv * TD / 2 + 4, side[1] + 15 };
                }
            },

            ["E"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.SideView,
                CalculatePosition = (wedge, drawing) =>
                {
                    var side = drawing.ViewPositions[Constants.SideView].GetValues(Unit.Millimeter);
                    var TD = wedge.Dimensions["TD"].GetValue(Unit.Millimeter);
                    var ssv = drawing.ViewScales[Constants.SideView].GetValue(Unit.Millimeter);
                    return new[] { side[0] + ssv * TD / 2 + 6, side[1] - 68 };
                }
            },
            ["FX"] = new DimensionLayoutRule
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
            ["F"] = new DimensionLayoutRule
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

                    return new[] { secv_x, secv_y - 10 };
                }
            },
            ["FL"] = new DimensionLayoutRule
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
            ["FR"] = new DimensionLayoutRule
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
            ["BR"] = new DimensionLayoutRule
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

    private static Dictionary<string, DimensionLayoutRule> GetOverlayRules()
    {
        return new Dictionary<string, DimensionLayoutRule>
        {
            ["FR"] = new()
            {
                BasedOnView = Constants.SectionView,
                CalculatePosition = (wedge, drawing) => new[] { 10.0, 10.0 } // Replace with actual logic
            },
            ["BR"] = new()
            {
                BasedOnView = Constants.SectionView,
                CalculatePosition = (wedge, drawing) => new[] { 20.0, 10.0 } // Replace with actual logic
            },
            ["ISA"] = new DimensionLayoutRule
            {
                BasedOnView = Constants.OverlayDetailView,
                CalculatePosition = (wedge, drawing) => new[] { 310.77, 148.33 }
            },

            ["GA"] = new()
            {
                BasedOnView = Constants.OverlayDetailView,
                CalculatePosition = (wedge, drawing) => new[] { 262.58, 148.33 } // Replace with actual logic
            },
            ["VW"] = new()
            {
                BasedOnView = Constants.FrontView,
                CalculatePosition = (wedge, drawing) => new[] { 50.0, 50.0 } // Replace with actual logic
            },
            ["E"] = new()
            {
                BasedOnView = Constants.SideView,
                CalculatePosition = (wedge, drawing) => new[] { 60.0, 60.0 } // Replace with actual logic
            },
            ["X"] = new()
            {
                BasedOnView = Constants.SideView,
                CalculatePosition = (wedge, drawing) => new[] { 70.0, 70.0 } // Replace with actual logic
            },
            ["TDF"] = new()
            {
                BasedOnView = Constants.OverlayTopView,
                CalculatePosition = (wedge, drawing) => new[] { 408.0, 18.5 } // Replace with actual logic
            },
            ["FX"] = new()
            {
                BasedOnView = Constants.OverlaySideView,
                CalculatePosition = (wedge, drawing) => new[] { 172.0, 22.0} // Replace with actual logic
            },
            ["D3"] = new()
            {
                BasedOnView = Constants.OverlaySideView,
                CalculatePosition = (wedge, drawing) => new[] { 270.0, 10.0 } // Replace with actual logic
            },
            ["FA"] = new()
            {
                BasedOnView = Constants.OverlaySideView,
                CalculatePosition = (wedge, drawing) => new[] { 370.332, 60.0 } // Replace with actual logic
            },
            ["BA"] = new()
            {
                BasedOnView = Constants.OverlaySideView,
                CalculatePosition = (wedge, drawing) => new[] { 345.0, 26.41 } // Replace with actual logic
            },
            ["E"] = new()
            {
                BasedOnView = Constants.OverlaySideView,
                CalculatePosition = (wedge, drawing) => new[] { 177.8, 49.21 } // Replace with actual logic
            },
        };
    }
}
