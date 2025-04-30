using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services;

public class ViewService : IViewService
{
    private View _swView;
    private ModelDoc2 _model;
    private DrawingDoc _drawingDoc;
    private string _viewName;
    private double _scaling;
    private bool _status;

    public ViewService(string viewName, ref ModelDoc2 model)
    {
        _viewName = viewName;
        _model = model;

        _status = _model.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
        if (!_status)
        {
            Logger.Warn($"Equation file not found: {viewName}");
        }

        _drawingDoc = _model as DrawingDoc;
        if (_drawingDoc == null)
        {
            Logger.Error("Model is not a DrawingDoc.");
            return;
        }

        _status = _drawingDoc.ActivateView(viewName);
        _swView = (View)_drawingDoc.ActiveDrawingView;

        if (_swView == null)
        {
            Logger.Warn($"Failed to get ActiveDrawingView after activating {viewName}");
        }
        else
        {
            _scaling = _swView.ScaleDecimal;
        }
    }

    public bool SetViewScale(double scale)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot set scale. View is null.");
            return false;
        }

        try
        {
            _scaling = scale;
            _swView.ScaleDecimal = scale;
            Logger.Info($"Set view scale to {scale:F3}.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to set view scale: {ex.Message}");
            return false;
        }
    }

    public bool SetViewPosition(DataStorage position)
    {
        if (_swView == null || position == null)
        {
            Logger.Warn("Cannot set position. View or Position is null.");
            return false;
        }

        try
        {
            _swView.Position = position.GetValues(Unit.Meter);
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to set view position: {ex.Message}");
            return false;
        }
    }

    public bool CreateFixedCenterline(DynamicDataContainer wedgeDimensions, DrawingData drawData)
    {
        try
        {
            string viewName = _swView.Name.ToLower();
            double scale = _swView.ScaleDecimal;

            double ScaleOffset(double mm) => mm / 1000 / scale;

            double[] GetCenterlineCoordinates()
            {
                switch (viewName)
                {
                    case "front_view":
                        double tlFront = wedgeDimensions["TL"].GetValue(Unit.Meter);
                        return new[]
                        {
                        0.0,
                        tlFront / 2 + ScaleOffset(2),
                        0.0,
                        -tlFront / 2 - ScaleOffset(2)
                    };

                    case "side_view":
                        double tlSide = wedgeDimensions["TL"].GetValue(Unit.Meter);
                        double tdf = wedgeDimensions["TDF"].GetValue(Unit.Meter);
                        double td = wedgeDimensions["TD"].GetValue(Unit.Meter);
                        double offset = (tdf - td) / 2;
                        return new[]
                        {
                        offset,
                        tlSide / 2 + ScaleOffset(2),
                        offset,
                        -tlSide / 2 - ScaleOffset(2)
                    };

                    case "detail_view":
                        double tlDetail = wedgeDimensions["TL"].GetValue(Unit.Meter);
                        return new[]
                        {
                        0.0,
                        0.0,
                        0.0,
                        -tlDetail / 2 + ScaleOffset(5)
                    };

                    default:
                        return null;
                }
            }

            double[] pos = GetCenterlineCoordinates();
            if (pos == null) return false;

            SketchSegment line = _model.SketchManager.CreateCenterLine(pos[0], pos[1], 0.0, pos[2], pos[3], 0.0);
            line.Layer = "FORMAT";
            line.GetSketch().RelationManager.AddRelation(new[] { line }, (int)swConstraintType_e.swConstraintType_FIXED);

            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool CreateFixedCentermark(DynamicDataContainer wedgeDimensions, DrawingData drawData)
    {
        if (_swView == null || _model == null)
        {
            Logger.Warn("Cannot create centermark. Model or View is null.");
            return false;
        }

        try
        {
            double scale = _swView.ScaleDecimal;
            double ScaleOffset(double mm) => mm / 1000 / scale;

            double td = wedgeDimensions["TD"].GetValue(Unit.Meter);
            double tdf = wedgeDimensions["TDF"].GetValue(Unit.Meter);
            double offset = (tdf - td) / 2;

            var centerlines = new[]
            {
            new { Start = new[] { offset, td/2 + ScaleOffset(2) }, End = new[] { offset, -td/2 - ScaleOffset(2) } },
            new { Start = new[] { offset + td/2 + ScaleOffset(2), 0.0 }, End = new[] { offset - td/2 - ScaleOffset(2), 0.0 } }
        };

            foreach (var linePoints in centerlines)
            {
                SketchSegment line = _model.SketchManager.CreateCenterLine(
                    linePoints.Start[0], linePoints.Start[1], 0,
                    linePoints.End[0], linePoints.End[1], 0);
                line.Layer = "FORMAT";
                line.GetSketch().RelationManager.AddRelation(new[] { line }, (int)swConstraintType_e.swConstraintType_FIXED);
            }

            Logger.Success("Created centermark cross-lines.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error creating centermark: {ex.Message}");
            return false;
        }
    }


    public bool SetBreaklinePosition(DynamicDataContainer wedgeDimensions, DrawingData drawData)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot set breakline position. View is null.");
            return false;
        }

        try
        {
            double scale = _swView.ScaleDecimal;
            double tl = wedgeDimensions["TL"].GetValue(Unit.Meter);
            BreakLine breakLine = _swView.IGetBreakLines(_swView.GetBreakLineCount2(out _));

            if (breakLine == null)
            {
                Logger.Warn("No breakline object found.");
                return false;
            }

            string viewName = _swView.Name.ToLower();
            (double lower, double upper, bool isDetail) = viewName switch
            {
                "front_view" => (drawData.BreaklineData["Front_viewLowerPartLength"].GetValue(Unit.Meter), drawData.BreaklineData["Front_viewLowerPartLength"].GetValue(Unit.Meter), false),
                "side_view" => (drawData.BreaklineData["Side_viewLowerPartLength"].GetValue(Unit.Meter), drawData.BreaklineData["Side_viewLowerPartLength"].GetValue(Unit.Meter), false),
                "detail_view" => (drawData.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Meter), tl / 2, true),
                _ => (0, 0, false)
            };

            double[] breaklinePos = isDetail
                ? new[] { lower - scale * tl / 2, scale * upper }
                : new[] { -tl * scale / 2 + lower, tl * scale / 2 - upper };

            bool result = breakLine.SetPosition(breaklinePos[0], breaklinePos[1]);

            if (isDetail)
            {
                _model.Extension.SelectByID2("TL@Detail_View", "DIMENSION", 0, 0, 0, false, 0, null, 0);
                bool deletion = _model.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Advanced);
                result &= deletion;
            }

            Logger.Info($"Breakline position set successfully in {viewName} view.");
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error setting breakline position: {ex.Message}");
            return false;
        }
    }


    public bool SetBreakLineGap(double gap)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot set breakline gap. View is null.");
            return false;
        }

        try
        {
            _swView.BreakLineGap = gap;
            Logger.Info($"Set breakline gap to {gap:F3} meters.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to set breakline gap: {ex.Message}");
            return false;
        }
    }

    public string CreateSectionView(
    IViewService parentView,
    DataStorage position,
    SketchSegment sketchSegment,
    DynamicDataContainer wedgeDimensions,
    DrawingData drawData)
    {
        if (_model == null || sketchSegment == null)
        {
            Logger.Warn("Cannot create section view. Model or SketchSegment is null.");
            return null;
        }

        try
        {
            Logger.Info("Starting section view creation...");

            _model.ClearSelection2(true);
            bool selected = sketchSegment.Select4(false, null);

            if (!selected)
            {
                Logger.Warn("Failed to select cutting sketch segment.");
                return null;
            }

            var drawingDoc = (DrawingDoc)_model;
            var view = drawingDoc.CreateSectionViewAt5(
                position.GetValues(Unit.Meter)[0],
                position.GetValues(Unit.Meter)[1],
                0.0,
                "",
                (int)swCreateSectionViewAtOptions_e.swCreateSectionView_ChangeDirection,
                null,
                0.01
            );

            if (view is null)
            {
                Logger.Warn("Section view creation returned null.");
                return null;
            }

            if (view.GetSection() is DrSection swSection)
            {
                swSection.SetAutoHatch(true);
                swSection.SetLabel2("Section_view");
            }

            string newName = view.Name;
            Logger.Success($"Section view created successfully: {newName}");
            return newName;
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception during section view creation: {ex.Message}");
            return null;
        }
    }


    public bool InsertModelDimensioning()
    {
        if (_swView == null || _model == null)
        {
            Logger.Warn("Cannot insert model dimensioning. Model or View is null.");
            return false;
        }

        try
        {
            Logger.Info($"Inserting model dimensions for view: {_swView.Name}");

            // Rebuild model and referenced part to ensure dimensions are up to date
            _model.ForceRebuild3(false);

            if (!_model.Extension.SelectByID2(_swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0))
            {
                Logger.Warn($"Failed to select drawing view: {_swView.Name}");
                return false;
            }

            if (_model is DrawingDoc drawingDoc)
            {
                drawingDoc.ActivateView(_swView.Name);
                _swView = (View)drawingDoc.ActiveDrawingView;

                var refModel = _swView?.ReferencedDocument;

                if (refModel == null)
                {
                    Logger.Warn("Referenced model is null. Cannot insert dimensions.");
                    return false;
                }

                // Rebuild and fit referenced model before inserting annotations
                refModel.ForceRebuild3(false);
                refModel.ViewZoomtofit2();

                // Insert all marked model dimensions
                object inserted = drawingDoc.InsertModelAnnotations3(
                    (int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel,
                    (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing,
                    false, true, false, false);

                if (inserted is object[] dimensions && dimensions.Length > 0)
                {
                    Logger.Success($"Inserted {dimensions.Length} model dimensions into view: {_swView.Name}");
                    return true;
                }

                Logger.Warn($"No model dimensions were inserted into view: {_swView.Name}");
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



    public bool SetPositionAndNameDimensioning(DynamicDataContainer wedgeDimensions, DynamicDimensioningContainer drawDimensions, Dictionary<string, string> dimensionTypes)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot set position and name. View is null.");
            return false;
        }

        try
        {
            DisplayDimension swDispDim = _swView.GetFirstDisplayDimension5();

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
                        SetAnnotationPositionAndName(swAnn, drawDimensions[dimKey].Position, dimKey);
                        break;
                    }
                    else if (selector == "SelectByValue" &&
                             wedgeDimensions.GetAll().TryGetValue(dimKey, out var modelValue))
                    {
                        double modelVal = modelValue.GetValue(Unit.Millimeter);
                        double dimVal = (double)swDim.GetSystemValue3((int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, "");

                        if (Math.Abs(modelVal - dimVal) < 1e-4)
                        {
                            SetAnnotationPositionAndName(swAnn, drawDimensions[dimKey].Position, dimKey);
                            break;
                        }
                    }
                }

                swDispDim = (DisplayDimension)swDispDim.GetNext3();
            }

            Logger.Success("Finished setting dimension positions and names.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error during SetPositionAndNameDimensioning: {ex.Message}");
            return false;
        }
    }




    private void SetAnnotationPositionAndName(Annotation swAnn, DataStorage pos, string name)
    {
        if (swAnn == null || pos == null)
        {
            Logger.Warn($"Cannot set annotation position for {name}: Missing Annotation or Position.");
            return;
        }

        double[] coords = pos.GetValues(Unit.Meter);

        if (coords.Length < 2)
        {
            Logger.Warn($"Invalid position array for dimension {name}");
            return;
        }

        try
        {
            swAnn.SetPosition2(coords[0], coords[1], 0.0);
            swAnn.Layer = "FORMAT";
            swAnn.SetName(name);

            Logger.Info($"Moved and renamed annotation '{name}' to ({coords[0]:F4}, {coords[1]:F4}) meters.");
        }
        catch (Exception ex)
        {
            Logger.Error($"Error setting annotation '{name}': {ex.Message}");
        }
    }


    public bool SetPositionAndLabelDatumFeature(DynamicDataContainer wedgeDimensions, DynamicDimensioningContainer drawDimensions, string label)
    {
        try
        {
            DatumTag datumTag = _swView.GetFirstDatumTag() as DatumTag;
            if (datumTag == null) return false;

            var ann = datumTag.GetAnnotation() as Annotation;
            double val = wedgeDimensions["SymmetryTolerance"].GetValue(Unit.Millimeter);

            if (val == 0.0 || double.IsNaN(val))
            {
                ann.Visible = (int)swAnnotationVisibilityState_e.swAnnotationHidden;
            }
            else
            {
                datumTag.SetLabel(label);
                var pos = drawDimensions["DatumFeature"].Position;
                ann.SetPosition2(pos.GetValues(Unit.Meter)[0], pos.GetValues(Unit.Meter)[1], 0.0);
            }

            return true;
        }
        catch
        {
            return false;
        }
    }
    public bool SetPositionAndValuesAndLabelGeometricTolerance(DynamicDataContainer wedgeDimensions, DynamicDimensioningContainer drawDimensions, string label)
    {
        try
        {
            Gtol gtol = _swView.GetFirstGTOL() as Gtol;
            if (gtol == null) return false;

            double symTol = wedgeDimensions["SymmetryTolerance"].GetValue(Unit.Millimeter);
            var ann = gtol.GetAnnotation() as Annotation;

            if (symTol == 0.0 || double.IsNaN(symTol))
            {
                ann.Visible = (int)swAnnotationVisibilityState_e.swAnnotationHidden;
            }
            else
            {
                var pos = drawDimensions["GeometricTolerance"].Position;
                gtol.SetPosition(pos.GetValues(Unit.Meter)[0], pos.GetValues(Unit.Meter)[1], 0.0);

                string tolInInch = Math.Round(wedgeDimensions["SymmetryTolerance"].GetValue(Unit.Millimeter), 4).ToString("0.0000");
                string tolInMm = "[" + wedgeDimensions["SymmetryTolerance"].GetValue(Unit.Millimeter).ToString("0.###") + "]";

                bool result = gtol.SetFrameValues2(1, tolInInch, "", tolInMm, label, "");
                return result;
            }

            return true;
        }
        catch
        {
            return false;
        }
    }
    public void ReactivateView(ref ModelDoc2 swModel)
    {
        if (swModel == null)
        {
            Logger.Warn("Cannot reactivate view. Model is null.");
            return;
        }

        try
        {
            string viewName = _swView?.Name ?? _viewName;

            if (string.IsNullOrWhiteSpace(viewName))
            {
                Logger.Warn("View name is invalid or empty.");
                return;
            }

            bool selected = swModel.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

            if (!selected)
            {
                Logger.Warn($"Failed to select view '{viewName}' for reactivation.");
                return;
            }

            if (swModel is DrawingDoc drawingDoc)
            {
                drawingDoc.ActivateView(viewName);
                _swView = (View)drawingDoc.ActiveDrawingView;

                if (_swView != null)
                    Logger.Info($"Reactivated view successfully: {_swView.GetName2()}");
                else
                    Logger.Warn($"Activated view '{viewName}' but could not retrieve active drawing view.");
            }
            else
            {
                Logger.Warn("Model is not a DrawingDoc during reactivation.");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception during ReactivateView: {ex.Message}");
        }
    }

    public double[] GetPosition()
    {
        if (_swView?.Position is double[] position && position.Length >= 2)
        {
            Logger.Info($"Section View Position: X = {position[0]:F4} m, Y = {position[1]:F4} m");
            return position;
        }

        Logger.Warn("Section view position is null or invalid. Returning (0,0).");
        return new double[] { 0.0, 0.0 };
    }


    public Dictionary<string, double[]> GetDefaultModelDimensionPositions()
    {
        var dimensionPositions = new Dictionary<string, double[]>();

        if (_swView == null)
        {
            Logger.Warn("Cannot capture dimensions. View is null.");
            return dimensionPositions;
        }

        try
        {
            DisplayDimension swDispDim = _swView.GetFirstDisplayDimension5();

            while (swDispDim != null)
            {
                var swAnn = swDispDim.GetAnnotation() as Annotation;
                var swDim = swDispDim.GetDimension2(0);

                if (swAnn != null && swDim != null)
                {
                    double[] pos = (double[])swAnn.GetPosition();

                    if (pos is { Length: >= 2 })
                    {
                        dimensionPositions[swDim.Name] = new[] { pos[0], pos[1] };
                        Logger.Info($"Captured dimension '{swDim.Name}': ({pos[0]:F4}, {pos[1]:F4}) meters.");
                    }
                }

                swDispDim = (DisplayDimension)swDispDim.GetNext3();
            }

            return dimensionPositions;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error during capturing default dimension positions: {ex.Message}");
            return dimensionPositions;
        }
    }


}