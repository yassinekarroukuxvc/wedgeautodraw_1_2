using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Core.Interfaces;

public interface IViewService
{
    bool SetViewScale(double scale);
    bool SetViewPosition(DataStorage position);
    bool CreateFixedCenterline(DynamicDataContainer wedgeDimensions, DrawingData drawData);
    bool CreateFixedCentermark(DynamicDataContainer wedgeDimensions, DrawingData drawData);
    bool SetBreaklinePosition(DynamicDataContainer wedgeDimensions, DrawingData drawData);
    bool SetBreakLineGap(double gap);
    public bool CreateSectionView(IViewService parentView, DataStorage position, SketchSegment sketchSegment, DynamicDataContainer wedgeDimensions, DrawingData drawData);
    bool InsertModelDimensioning();
    bool SetPositionAndNameDimensioning(DynamicDataContainer wedgeDimensions, DynamicDimensioningContainer drawDimensions, Dictionary<string, string> dimensionTypes);
    bool SetPositionAndLabelDatumFeature(DynamicDataContainer wedgeDimensions, DynamicDimensioningContainer drawDimensions, string label);
    bool SetPositionAndValuesAndLabelGeometricTolerance(DynamicDataContainer wedgeDimensions, DynamicDimensioningContainer drawDimensions, string label);
    void ReactivateView(ref ModelDoc2 swModel);
}
