using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;

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
        _drawingDoc = (DrawingDoc)_model;
        _status = _drawingDoc.ActivateView(viewName);
        _swView = (View)_drawingDoc.ActiveDrawingView;
        _scaling = _swView.ScaleDecimal;
    }
    public bool SetViewScale(double scale)
    {
        try
        {
            _scaling = scale;
            _swView.ScaleDecimal = scale;
            return true;
        }
        catch
        {
            return false;
        }
    }
    public bool SetViewPosition(DataStorage position)
    {
        try
        {
            var pos = position.GetValues(Unit.Meter);
            _swView.Position = pos;
            return true;
        }
        catch
        {
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
        try
        {
            double scale = _swView.ScaleDecimal;
            double ScaleOffset(double mm) => mm / 1000 / scale;

            double td = wedgeDimensions["TD"].GetValue(Unit.Meter);
            double tdf = wedgeDimensions["TDF"].GetValue(Unit.Meter);
            double offset = (tdf - td) / 2;

            double[][] centerlinePoints =
            {
            new[] { offset, td / 2 + ScaleOffset(2), offset, -td / 2 - ScaleOffset(2) },
            new[] { offset + td / 2 + ScaleOffset(2), 0.0, offset - td / 2 - ScaleOffset(2), 0.0 }
        };

            foreach (var pos in centerlinePoints)
            {
                SketchSegment line = _model.SketchManager.CreateCenterLine(pos[0], pos[1], 0.0, pos[2], pos[3], 0.0);
                line.Layer = "FORMAT";
                line.GetSketch().RelationManager.AddRelation(new[] { line }, (int)swConstraintType_e.swConstraintType_FIXED);
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool SetBreaklinePosition(DynamicDataContainer wedgeDimensions, DrawingData drawData)
    {
        try
        {
            string viewName = _swView.Name.ToLower();
            double scale = _swView.ScaleDecimal;
            double tl = wedgeDimensions["TL"].GetValue(Unit.Meter);
            BreakLine breakLine = _swView.IGetBreakLines(_swView.GetBreakLineCount2(out _));

            double lower = 0, upper = 0;
            bool isDetailView = false;

            switch (viewName)
            {
                case "front_view":
                    lower = drawData.BreaklineData["Front_viewLowerPartLength"].GetValue(Unit.Meter);
                    upper = drawData.BreaklineData["Front_viewLowerPartLength"].GetValue(Unit.Meter);
                    break;

                case "side_view":
                    lower = drawData.BreaklineData["Side_viewLowerPartLength"].GetValue(Unit.Meter);
                    upper = drawData.BreaklineData["Side_viewLowerPartLength"].GetValue(Unit.Meter);
                    break;

                case "detail_view":
                    lower = drawData.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Meter);
                    upper = tl / 2;
                    isDetailView = true;
                    break;

                default:
                    return false;
            }

            double[] breaklinePos = isDetailView
                ? new[]
                {
                lower - scale * tl / 2,
                scale * upper
                }
                : new[]
                {
                -tl * scale / 2 + lower,
                tl * scale / 2 - upper
                };

            bool result = breakLine.SetPosition(breaklinePos[0], breaklinePos[1]);

            if (isDetailView)
            {
                _model.Extension.SelectByID2("TL@Detail_View", "DIMENSION", 0, 0, 0, false, 0, null, 0);
                bool deletion = _model.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Advanced);
                result &= deletion;
            }

            return result;
        }
        catch
        {
            return false;
        }
    }

    public bool SetBreakLineGap(double gap)
    {
        try
        {
            _swView.BreakLineGap = gap;
            return true;
        }
        catch
        {
            return false;
        }
    }
    public bool CreateSectionView(IViewService parentView, DataStorage position, SketchSegment sketchSegment, DynamicDataContainer wedgeDimensions, DrawingData drawData)
    {
        try
        {
            Console.WriteLine("🔄 Starting section view creation...");

            // Access drawing and model from current context
            var drawingDoc = (DrawingDoc)_model;

            // Clear any existing selections and select the sketch segment
            _model.ClearSelection2(true);
            bool selected = sketchSegment.Select4(false, null);

            if (!selected)
            {
                Console.WriteLine("❌ Failed to select sketch segment.");
                return false;
            }

            // Create the section view
            var sectionView = drawingDoc.CreateSectionViewAt5(
                position.GetValues(Unit.Meter)[0],
                position.GetValues(Unit.Meter)[1],
                0.0,
                "", // label
                (int)swCreateSectionViewAtOptions_e.swCreateSectionView_ChangeDirection,
                null, // use selection
                0.01 // depth
            );

            if (sectionView == null)
            {
                Console.WriteLine("❌ Section view creation failed.");
                return false;
            }

            var swSectionView = (DrSection)sectionView.GetSection();
            swSectionView.SetAutoHatch(true);
            swSectionView.SetLabel2("");
            swSectionView.CutSurfaceBodies = false;
            swSectionView.CuttingLineShoulders = false;
            swSectionView.DisplaySurfaceBodies = false;
            swSectionView.ExcludeSliceSectionBodies = false;

            Console.WriteLine("✅ Section view created successfully.");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine("❌ Exception during section view creation: " + ex.Message);
            return false;
        }
    }
    public bool InsertModelDimensioning()
    {
        try
        {
            var drawingDoc = (DrawingDoc)_swView.GetType().InvokeMember("Model", System.Reflection.BindingFlags.GetProperty, null, _swView, null);
            object result = drawingDoc.InsertModelAnnotations3(
                (int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel,
                (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing,
                false, true, false, false
            );

            return result is object[] inserted && inserted.Length > 0;
        }
        catch
        {
            return false;
        }
    }
    public bool SetPositionAndNameDimensioning(DynamicDataContainer wedgeDimensions, DynamicDimensioningContainer drawDimensions, Dictionary<string, string> dimensionTypes)
    {
        try
        {
            DisplayDimension swDispDim = _swView.GetFirstDisplayDimension5();

            while (swDispDim != null)
            {
                var swAnn = swDispDim.GetAnnotation() as Annotation;
                var swDim = swDispDim.GetDimension2(0);

                foreach (var (dimKey, selector) in dimensionTypes)
                {
                    if (selector == "SelectByName" && swDim.Name == dimKey)
                    {
                        SetAnnotationPositionAndName(swAnn, drawDimensions[dimKey].Position, dimKey);
                    }
                    else if (selector == "SelectByValue")
                    {
                        if (wedgeDimensions.GetAll().TryGetValue(dimKey, out var modelValue))
                        {
                            double modelVal = modelValue.GetValue(Unit.Millimeter);
                            double dimVal = (double)swDim.GetSystemValue3((int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, "");

                            if (Math.Abs(modelVal - dimVal) < 1e-4)
                            {
                                SetAnnotationPositionAndName(swAnn, drawDimensions[dimKey].Position, dimKey);
                            }
                        }
                    }
                }

                swDispDim = (DisplayDimension)swDispDim.GetNext3();
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private void SetAnnotationPositionAndName(Annotation swAnn, DataStorage pos, string name)
    {
        double[] coords = pos.GetValues(Unit.Meter);
        swAnn.SetPosition2(coords[0], coords[1], 0.0);
        swAnn.Layer = "FORMAT";
        swAnn.SetName(name);
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

                string tolInInch = Math.Round(wedgeDimensions["SymmetryTolerance"].GetValue(Unit.Inch), 4).ToString("0.0000");
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
        try
        {
            if (swModel == null) return;

            string viewName = _swView.Name;
            bool selected = swModel.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

            if (selected && swModel is DrawingDoc drawingDoc)
            {
                drawingDoc.ActivateView(viewName);
                _swView = (View)drawingDoc.ActiveDrawingView;
            }
        }
        catch
        {
            Console.WriteLine("Error while Reactivating the View");
        }
    }


}
