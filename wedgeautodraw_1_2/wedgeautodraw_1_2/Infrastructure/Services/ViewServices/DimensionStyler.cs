using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services.ViewServices;

public class DimensionStyler
{
    private readonly ModelDoc2 _model;

    public DimensionStyler(ModelDoc2 model)
    {
        _model = model;
    }

    public bool Apply(View swView, NamedDimensionValues wedgeDimensions, NamedDimensionAnnotations drawDimensions, Dictionary<string, string> dimensionTypes)
    {
        if (swView == null)
        {
            Logger.Warn("Cannot apply dimension styles. View is null.");
            return false;
        }

        try
        {
            DisplayDimension swDispDim = swView.GetFirstDisplayDimension5();

            while (swDispDim != null)
            {
                var swAnn = swDispDim.GetAnnotation() as Annotation;
                var swDim = swDispDim.GetDimension2(0);

                if (swAnn == null || swDim == null)
                {
                    swDispDim = (DisplayDimension)swDispDim.GetNext3();
                    continue;
                }

                foreach (var (dimKey, selector) in dimensionTypes)
                {
                    if (selector == "SelectByName" && swDim.Name == dimKey)
                    {
                        UpdateAnnotation(swAnn, drawDimensions[dimKey].Position, dimKey);
                        Style(swDispDim, dimKey, wedgeDimensions, drawDimensions);
                        break;
                    }
                    else if (selector == "SelectByValue" && wedgeDimensions.TryGet(dimKey, out var modelValue))
                    {
                        double modelVal = modelValue.GetValue(Unit.Millimeter);
                        double dimVal = (double)swDim.GetSystemValue3((int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, "");

                        if (Math.Abs(modelVal - dimVal) < 1e-4)
                        {
                            UpdateAnnotation(swAnn, drawDimensions[dimKey].Position, dimKey);
                            Style(swDispDim, dimKey, wedgeDimensions, drawDimensions);
                            break;
                        }
                    }
                }

                swDispDim = (DisplayDimension)swDispDim.GetNext3();
            }

            Logger.Success("Finished applying dimension positions and styles.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error applying dimension styles: {ex.Message}");
            return false;
        }
    }

    private void UpdateAnnotation(Annotation swAnn, DataStorage pos, string name, double verticalOffset = 0.0)
    {
        if (swAnn == null || pos == null)
        {
            Logger.Warn($"Cannot update annotation for {name}: null reference.");
            return;
        }

        double[] coords = pos.GetValues(Unit.Meter);
        if (coords.Length < 2)
        {
            Logger.Warn($"Invalid position for dimension {name}.");
            return;
        }

        try
        {
            swAnn.SetPosition2(coords[0], coords[1] - verticalOffset, 0.0);
            swAnn.Layer = "FORMAT";
            swAnn.SetName(name);
            Logger.Info($"Updated annotation '{name}' to ({coords[0]:F4}, {coords[1] - verticalOffset:F4}) m.");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to update annotation '{name}': {ex.Message}");
        }
    }

    private void Style(DisplayDimension swDispDim, string dimKey, NamedDimensionValues wedgeDimensions, NamedDimensionAnnotations drawDimensions)
    {
        try
        {
            if (dimKey != "ISA")
                swDispDim.CenterText = true;

            if (dimKey == "VW")
            {
                swDispDim.CenterText = true;
                Logger.Success("VW is centered");
            }

            if (dimKey == "FL")
            {
                swDispDim.ArrowSide = 0;
                swDispDim.WitnessVisibility = 0;
                Logger.Success("FL has arrow outside");
            }

            if (dimKey == "GA")
            {
                swDispDim.ArrowSide = 0;

                if (wedgeDimensions.TryGet("W", out var wVal) && wVal.GetValue(Unit.Millimeter) < 0.3)
                {
                    swDispDim.OffsetText = true;
                    swDispDim.CenterText = true;

                    if (drawDimensions.ContainsKey("GA") && drawDimensions["GA"].Position != null)
                    {
                        var swAnn = swDispDim.GetAnnotation() as Annotation;
                        UpdateAnnotation(swAnn, drawDimensions["GA"].Position, "GA", 0.005); // 5mm offset
                        Logger.Success("GA position reapplied with 5mm vertical offset due to W < 0.9mm.");
                    }
                }
            }
            if (dimKey == "E")
            {
                swDispDim.CenterText = true;

                if (wedgeDimensions.TryGet("E", out var wVal) && wVal.GetValue(Unit.Millimeter) < 0.8)
                {
                    swDispDim.OffsetText = true;

                    if (drawDimensions.ContainsKey("E") && drawDimensions["E"].Position != null)
                    {
                        var swAnn = swDispDim.GetAnnotation() as Annotation;
                        UpdateAnnotation(swAnn, drawDimensions["E"].Position, "E", 0.010); // 5mm offset
                        Logger.Success("E position reapplied with 5mm vertical offset due to E < 0.8mm.");
                    }
                }
            }

            if (dimKey == "GD")
            {
                swDispDim.WitnessVisibility = 0;
            }
            if (dimKey == "FA" || dimKey == "BA")
            {
                swDispDim.CenterText = true;
            }

            if (dimKey is "FR" or "BR")
            {
                swDispDim.ArrowSide = 2;
                swDispDim.WitnessVisibility = 2;
                swDispDim.ExtensionLineExtendsFromCenterOfSet = false;
                swDispDim.MaxWitnessLineLength = 0;
                swDispDim.SetExtensionLineAsCenterline(1, false);
            }

            _model.GraphicsRedraw2();
            Logger.Info($"Styled dimension '{dimKey}'.");
        }
        catch (Exception ex)
        {
            Logger.Warn($"Failed to style '{dimKey}': {ex.Message}");
        }
    }
}
