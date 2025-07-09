using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Core.Interfaces;

public interface IViewService
{
    IView GetRawView();
    bool SetViewScale(double scale);
    bool SetViewPosition(DataStorage position);
    bool CreateCenterline(NamedDimensionValues wedgeDimensions, DrawingData drawData);
    bool CreateCentermark(NamedDimensionValues wedgeDimensions, DrawingData drawData);
    bool SetBreaklinePosition(NamedDimensionValues wedgeDimensions, DrawingData drawData);
    bool SetBreakLineGap(double gap);
    public string CreateSectionView(IViewService parentView, DataStorage position, SketchSegment sketchSegment, NamedDimensionValues wedgeDimensions, DrawingData drawData);
    bool InsertModelDimensioning(DrawingType drawingType);
    bool ApplyDimensionPositionsAndNames(NamedDimensionValues wedgeDimensions, NamedDimensionAnnotations drawDimensions, Dictionary<string, string> dimensionTypes, DrawingType drawingType);
    bool PlaceDatumFeatureLabel(NamedDimensionValues wedgeDimensions, NamedDimensionAnnotations drawDimensions, string label);
    bool PlaceGeometricToleranceFrame(NamedDimensionValues wedgeDimensions, NamedDimensionAnnotations drawDimensions, string label);
    void ReactivateView(ref ModelDoc2 swModel);
    bool DeleteAnnotationsByName(string[] annotationNames);
    public bool SetOverlayBreaklinePosition(NamedDimensionValues wedgeDimensions, DrawingData drawData);
    public bool CenterSectionViewVisuallyVertically(NamedDimensionValues wedgeDimensions);
    public void AlignViewHorizontally(bool isDetailView, double tlInMeters = 0);
    public void SetSketchDimensionValue(string dimensionName, double value);
    public void CenterViewVertically();
    public void PositionSideViewHorizontally(double tlInMeters);
    public bool CreateCenterlineAtViewCenter(bool isVertical = true, double lengthMm = 100.0);
    public void AlignTopViewNextToSideView(IView sideView, IView topView, double offsetMm = 30.0);
    public double GetSketchDimensionValue(string dimensionName);
    public bool SetDetailViewDynamicBreakline(NamedDimensionValues wedgeDimensions);
    public bool SetFrontSideViewBreakline(NamedDimensionValues wedgeDimensions);
}
