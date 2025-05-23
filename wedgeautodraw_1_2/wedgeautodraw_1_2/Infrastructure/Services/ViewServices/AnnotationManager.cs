using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services.ViewServices;

public class AnnotationManager
{
    private readonly View _swView;

    public AnnotationManager(View swView)
    {
        _swView = swView;
    }

    public bool PlaceDatumFeatureLabel(NamedDimensionValues wedgeDimensions, NamedDimensionAnnotations drawDimensions, string label)
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
        catch (Exception ex)
        {
            Logger.Error($"Failed to place datum feature label: {ex.Message}");
            return false;
        }
    }

    public bool PlaceGeometricToleranceFrame(NamedDimensionValues wedgeDimensions, NamedDimensionAnnotations drawDimensions, string label)
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

                string tolInInch = Math.Round(symTol, 4).ToString("0.0000");
                string tolInMm = "[" + symTol.ToString("0.###") + "]";

                bool result = gtol.SetFrameValues2(1, tolInInch, "", tolInMm, label, "");
                return result;
            }

            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to place GTOL frame: {ex.Message}");
            return false;
        }
    }
}
