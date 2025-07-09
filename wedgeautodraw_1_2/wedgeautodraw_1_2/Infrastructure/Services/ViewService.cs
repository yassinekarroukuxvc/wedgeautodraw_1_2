using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Spreadsheet;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Reflection.Emit;
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
    private ViewPositionManager _viewPositionManager;
    private CenterMarkLineManager _centerMarkLineManager;

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
        _annotationManager = new AnnotationManager(_swView,_model);
        _sectionViewCreator = new SectionViewCreator(_model);
        _viewPositionManager = new ViewPositionManager(_swView, _model,_drawingDoc);
        _centerMarkLineManager = new CenterMarkLineManager(_swView, _model);
    }
    public IView GetRawView()
    {
        return _swView;
    }
    public bool SetViewScale(double scale) => _scaler.SetScale(scale);

    public bool SetViewPosition(DataStorage position)
    => _viewPositionManager.SetViewPosition(position);

    public bool CreateCenterline(NamedDimensionValues wedgeDimensions, DrawingData drawData)
    => _centerMarkLineManager.CreateCenterline(wedgeDimensions, drawData);

    public bool CreateCentermark(NamedDimensionValues wedgeDimensions, DrawingData drawData)
    => _centerMarkLineManager.CreateCentermark(wedgeDimensions, drawData);

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

    public bool InsertModelDimensioning(DrawingType drawingType)
    => _dimensionStyler.Insert(_swView, drawingType);

    public bool ApplyDimensionPositionsAndNames(
        NamedDimensionValues wedgeDimensions,
        NamedDimensionAnnotations drawDimensions,
        Dictionary<string, string> dimensionTypes, DrawingType drawingType)
        => _dimensionStyler.Apply(_swView, wedgeDimensions, drawDimensions, dimensionTypes, drawingType);

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
    public bool SetOverlayBreaklinePosition(NamedDimensionValues wedgeDimensions, DrawingData drawData)
        => _breaklineHandler.SetOverlayBreaklinePosition(wedgeDimensions, drawData);
    public void CenterViewVertically()
    => _viewPositionManager.CenterViewVertically();

    public void AlignViewHorizontally(bool isDetailView, double tlInMeters = 0)
        => _viewPositionManager.CenterViewHorizontally(isDetailView, tlInMeters);
    public bool CenterSectionViewVisuallyVertically(NamedDimensionValues wedgeDimensions)
    => _viewPositionManager.CenterSectionViewVisuallyVertically(wedgeDimensions);
    public void SetSketchDimensionValue(string dimensionName, double value)
    {
        try
        {
            // Example: "D1@Sketch100"            // Example: "D1@Sketch100"
            bool selected = _model.Extension.SelectByID2(
                dimensionName,
                "DIMENSION",
                0, 0, 0,
                false, 0, null, 0
            );

            if (!selected)
            {
                Logger.Warn($"Failed to select dimension: {dimensionName}");
                return;
            }

            // Get selected DisplayDimension
            var selectionMgr = (ISelectionMgr)_model.SelectionManager;
            var dispDim = selectionMgr.GetSelectedObject6(1, -1) as DisplayDimension;

            if (dispDim == null)
            {
                Logger.Warn($"Selected object is not a DisplayDimension: {dimensionName}");
                return;
            }

            var dim = dispDim.GetDimension2(0);
            if (dim == null)
            {
                Logger.Warn($"Failed to get Dimension2 from DisplayDimension: {dimensionName}");
                return;
            }

            // Set the value (in METERS)
            dim.SystemValue = value; // value in meters!

            Logger.Success($"Set dimension '{dimensionName}' to {value} meters.");
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception while setting dimension '{dimensionName}': {ex.Message}");
        }
    }

    public double GetSketchDimensionValue(string dimensionName)
    {
        try
        {
            bool selected = _model.Extension.SelectByID2(
                dimensionName,
                "DIMENSION",
                0, 0, 0,
                false, 0, null, 0
            );

            if (!selected)
            {
                Logger.Warn($"Failed to select dimension: {dimensionName}");
                return double.NaN;
            }

            var selectionMgr = (ISelectionMgr)_model.SelectionManager;
            var dispDim = selectionMgr.GetSelectedObject6(1, -1) as DisplayDimension;

            if (dispDim == null)
            {
                Logger.Warn($"Selected object is not a DisplayDimension: {dimensionName}");
                return double.NaN;
            }

            var dim = dispDim.GetDimension2(0);
            if (dim == null)
            {
                Logger.Warn($"Failed to get Dimension2 from DisplayDimension: {dimensionName}");
                return double.NaN;
            }

            return dim.SystemValue;
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception while getting dimension '{dimensionName}': {ex.Message}");
            return double.NaN;
        }
    }

    public void PositionSideViewHorizontally(double tlInMeters)
        => _viewPositionManager.PositionSideViewHorizontally(tlInMeters);
    public bool CreateCenterlineAtViewCenter(bool isVertical = true, double lengthMm = 100.0)
        => _centerMarkLineManager.CreateCenterlineAtViewCenter(isVertical, lengthMm);
    public void AlignTopViewNextToSideView(IView sideView, IView topView, double offsetMm = 30.0)
        => _viewPositionManager.AlignTopViewNextToSideView(sideView, topView, offsetMm);

    public bool SetDetailViewDynamicBreakline(NamedDimensionValues wedgeDimensions)
        => _breaklineHandler.SetDetailViewDynamicBreakline(wedgeDimensions);
    public bool SetFrontSideViewBreakline(NamedDimensionValues wedgeDimensions)
        => _breaklineHandler.SetFrontSideViewBreakline(wedgeDimensions);
}
