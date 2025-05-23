using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Spreadsheet;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;
using wedgeautodraw_1_2.Infrastructure.Services.ViewServices;

namespace wedgeautodraw_1_2.Infrastructure.Services;

public class ViewService : IViewService
{
    private View _swView;
    private ModelDoc2 _model;
    private DrawingDoc _drawingDoc;
    private string _viewName;
    private bool _status;

    private ViewScaler _scaler;
    private BreaklineHandler _breaklineHandler;
    private DimensionStyler _dimensionStyler;
    private AnnotationManager _annotationManager;
    private SectionViewCreator _sectionViewCreator;

    public ViewService(string viewName, ref ModelDoc2 model)
    {
        _viewName = viewName;
        _model = model;

        _status = _model.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
        if (!_status)
            Logger.Warn($"View not found: {viewName}");

        _drawingDoc = _model as DrawingDoc;
        if (_drawingDoc == null)
        {
            Logger.Error("Model is not a DrawingDoc.");
            return;
        }

        _status = _drawingDoc.ActivateView(viewName);
        _swView = (View)_drawingDoc.ActiveDrawingView;

        if (_swView == null)
            Logger.Warn($"Failed to get ActiveDrawingView after activating {viewName}");

        _scaler = new ViewScaler(_swView);
        _breaklineHandler = new BreaklineHandler(_swView, _model);
        _dimensionStyler = new DimensionStyler(_model);
        _annotationManager = new AnnotationManager(_swView);
        _sectionViewCreator = new SectionViewCreator(_model);
    }

    public bool SetViewScale(double scale) => _scaler.SetScale(scale);

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

    public bool CreateCenterline(NamedDimensionValues wedgeDimensions, DrawingData drawData)
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
                        -tlDetail / 2 + ScaleOffset(10)
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

    public bool CreateCentermark(NamedDimensionValues wedgeDimensions, DrawingData drawData)
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

    public bool SetBreaklinePosition(NamedDimensionValues wedgeDimensions, DrawingData drawData)
        => _breaklineHandler.SetBreaklinePosition(wedgeDimensions, drawData);

    public bool SetBreakLineGap(double gap)
        => _breaklineHandler.SetBreaklineGap(gap);

    public string CreateSectionView(
        IViewService parentView,
        DataStorage position,
        SketchSegment sketchSegment,
        NamedDimensionValues wedgeDimensions,
        DrawingData drawData)
        => _sectionViewCreator.Create(position, sketchSegment, wedgeDimensions, drawData);

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

                refModel.ForceRebuild3(false);
                refModel.ViewZoomtofit2();

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

    public bool ApplyDimensionPositionsAndNames(
        NamedDimensionValues wedgeDimensions,
        NamedDimensionAnnotations drawDimensions,
        Dictionary<string, string> dimensionTypes)
        => _dimensionStyler.Apply(_swView, wedgeDimensions, drawDimensions, dimensionTypes);

    public bool PlaceDatumFeatureLabel(NamedDimensionValues wedgeDimensions, NamedDimensionAnnotations drawDimensions, string label)
        => _annotationManager.PlaceDatumFeatureLabel(wedgeDimensions, drawDimensions, label);

    public bool PlaceGeometricToleranceFrame(NamedDimensionValues wedgeDimensions, NamedDimensionAnnotations drawDimensions, string label)
        => _annotationManager.PlaceGeometricToleranceFrame(wedgeDimensions, drawDimensions, label);

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

        //Logger.Warn("Section view position is null or invalid. Returning (0,0).\");
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

    public bool RotateView(double angleInDegrees)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot rotate view. View is null.");
            return false;
        }

        try
        {
            double angleInRadians = angleInDegrees * Math.PI / 180.0;
            _swView.Angle = angleInRadians;
            _model.EditRebuild3();
            Logger.Info($"Rotated view '{_swView.Name}' by {angleInDegrees} degrees ({angleInRadians:F4} radians).");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to rotate view '{_swView.Name}': {ex.Message}");
            return false;
        }
    }
    public bool DeleteAnnotationsByName(string[] annotationNames)
    {
        if (_swView == null || _model == null)
        {
            Logger.Warn("Cannot delete annotations. View or model is null.");
            return false;
        }

        try
        {
            int deletedCount = 0;
            DisplayDimension swDispDim = _swView.GetFirstDisplayDimension5();

            while (swDispDim != null)
            {
                var swAnn = swDispDim.GetAnnotation() as Annotation;
                var swDim = swDispDim.GetDimension2(0);

                if (swAnn != null && swDim != null && annotationNames.Contains(swDim.Name))
                {
                    // Select the annotation
                    bool selected = swAnn.Select2(false, -1);
                    if (!selected)
                    {
                        Logger.Warn($"Failed to select annotation '{swDim.Name}' for deletion.");
                        swDispDim = (DisplayDimension)swDispDim.GetNext3();
                        continue;
                    }

                    // Delete it
                    bool deleted = _model.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                    if (deleted)
                    {
                        Logger.Info($"Deleted annotation '{swDim.Name}'.");
                        deletedCount++;
                    }
                    else
                    {
                        Logger.Warn($"Failed to delete annotation '{swDim.Name}'.");
                    }
                }

                swDispDim = (DisplayDimension)swDispDim.GetNext3();
            }

            Logger.Success($"Deleted {deletedCount} annotations.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error deleting annotations: {ex.Message}");
            return false;
        }
    }

}
