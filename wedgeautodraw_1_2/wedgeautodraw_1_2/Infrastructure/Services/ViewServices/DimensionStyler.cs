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
    public bool Insert(View swView)
    {
        if (swView == null || _model == null)
        {
            Logger.Warn("Cannot insert model dimensioning. Model or View is null.");
            return false;
        }

        try
        {
            Logger.Info($"Inserting model dimensions for view: {swView.Name}");
            _model.ForceRebuild3(false);

            if (!_model.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0))
            {
                Logger.Warn($"Failed to select drawing view: {swView.Name}");
                return false;
            }

            if (_model is DrawingDoc drawingDoc)
            {
                drawingDoc.ActivateView(swView.Name);
                swView = (View)drawingDoc.ActiveDrawingView;

                var refModel = swView?.ReferencedDocument;
                if (refModel == null)
                {
                    Logger.Warn("Referenced model is null. Cannot insert dimensions.");
                    return false;
                }

                refModel.ForceRebuild3(false);
                refModel.ViewZoomtofit2();

                object inserted = drawingDoc.InsertModelAnnotations3(
                    (int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel,
                    (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing,
                    false, true, false, false);

                if (inserted is object[] dimensions && dimensions.Length > 0)
                {
                    Logger.Success($"Inserted {dimensions.Length} model dimensions into view: {swView.Name}");
                    return true;
                }

                Logger.Warn($"No model dimensions were inserted into view: {swView.Name}");
                return false;
            }

            Logger.Warn("Active model is not a DrawingDoc.");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception during InsertModelDimensioning: {ex.Message}");
            return false;
        }
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
                        TryUpdateAnnotation(swAnn, dimKey, drawDimensions);
                        Style(swDispDim, dimKey, wedgeDimensions, drawDimensions);
                        break;
                    }
                    else if (selector == "SelectByValue" && wedgeDimensions.TryGet(dimKey, out var modelValue))
                    {
                        double modelVal = modelValue.GetValue(Unit.Millimeter);
                        double dimVal = (double)swDim.GetSystemValue3((int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, "");

                        if (Math.Abs(modelVal - dimVal) < 1e-4)
                        {
                            TryUpdateAnnotation(swAnn, dimKey, drawDimensions);
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

    private void TryUpdateAnnotation(Annotation swAnn, string dimKey, NamedDimensionAnnotations drawDimensions, double verticalOffset = 0.0)
    {
        if (swAnn == null || !drawDimensions.TryGet(dimKey, out var annotation) || annotation.Position == null)
        {
            Logger.Warn($"Cannot update annotation for {dimKey}: missing annotation or position.");
            return;
        }

        var coords = annotation.Position.GetValues(Unit.Meter);
        if (coords.Length < 2)
        {
            Logger.Warn($"Invalid position array for {dimKey}.");
            return;
        }

        try
        {
            swAnn.SetPosition2(coords[0], coords[1] - verticalOffset, 0.0);
            swAnn.Layer = "FORMAT";
            swAnn.SetName(dimKey);
            Logger.Info($"Updated annotation '{dimKey}' to ({coords[0]:F4}, {coords[1] - verticalOffset:F4}) m.");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to update annotation '{dimKey}': {ex.Message}");
        }
    }

    private void Style(DisplayDimension swDispDim, string dimKey, NamedDimensionValues wedgeDimensions, NamedDimensionAnnotations drawDimensions)
    {
        try
        {
            // Default: center text
            if (dimKey != "ISA")
                swDispDim.CenterText = true;

            switch (dimKey)
            {
                case "VW":
                    swDispDim.CenterText = true;
                    Logger.Success("VW is centered.");
                    break;

                case "FL":
                    swDispDim.ArrowSide = 0;
                    swDispDim.WitnessVisibility = 0;
                    Logger.Success("FL has arrow outside.");
                    break;

                case "GA":
                    swDispDim.ArrowSide = 0;
                    if (wedgeDimensions.TryGet("W", out var wVal) && wVal.GetValue(Unit.Millimeter) < 0.3)
                    {
                        swDispDim.OffsetText = true;
                        swDispDim.CenterText = true;
                        TryOffsetUpdate(swDispDim, "GA", drawDimensions, 0.005); // 5mm offset
                    }
                    break;

                case "E":
                    swDispDim.CenterText = true;
                    if (wedgeDimensions.TryGet("E", out var eVal) && eVal.GetValue(Unit.Millimeter) < 0.8)
                    {
                        swDispDim.OffsetText = true;
                        TryOffsetUpdate(swDispDim, "E", drawDimensions, 0.010); // 10mm offset
                    }
                    break;

                case "GD":
                    swDispDim.WitnessVisibility = 0;
                    break;

                case "FA":
                case "BA":
                    swDispDim.CenterText = true;
                    break;

                case "FR":
                case "BR":
                    swDispDim.ArrowSide = 2;
                    swDispDim.WitnessVisibility = 2;
                    swDispDim.ExtensionLineExtendsFromCenterOfSet = false;
                    swDispDim.MaxWitnessLineLength = 0;
                    swDispDim.SetExtensionLineAsCenterline(1, false);
                    break;
            }

            _model.GraphicsRedraw2();
            Logger.Info($"Styled dimension '{dimKey}'.");
        }
        catch (Exception ex)
        {
            Logger.Warn($"Failed to style '{dimKey}': {ex.Message}");
        }
    }

    private void TryOffsetUpdate(DisplayDimension swDispDim, string dimKey, NamedDimensionAnnotations drawDimensions, double offsetMeters)
    {
        if (!drawDimensions.TryGet(dimKey, out var dimAnn) || dimAnn.Position == null)
            return;

        var swAnn = swDispDim.GetAnnotation() as Annotation;
        TryUpdateAnnotation(swAnn, dimKey, drawDimensions, offsetMeters);
        Logger.Success($"{dimKey} position reapplied with offset of {offsetMeters * 1000:F1}mm.");
    }
}
