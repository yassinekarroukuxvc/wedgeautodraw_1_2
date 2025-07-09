using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services.ViewServices;

public class AnnotationManager
{
    private readonly View _swView;
    private readonly ModelDoc2 _model;

    public AnnotationManager(View swView, ModelDoc2 model)
    {
        _swView = swView;
        _model = model;
    }


    public bool PlaceDatumFeatureLabel(NamedDimensionValues wedgeDimensions, NamedDimensionAnnotations drawDimensions, string label)
    {
        try
        {
            if (_swView?.GetFirstDatumTag() is not DatumTag datumTag)
            {
                Logger.Warn("No DatumTag found in view.");
                return false;
            }

            if (datumTag.GetAnnotation() is not Annotation ann)
            {
                Logger.Warn("DatumTag annotation is null.");
                return false;
            }

            if (!wedgeDimensions.TryGet("SymmetryTolerance", out var tolStorage))
            {
                Logger.Warn("SymmetryTolerance not found.");
                ann.Visible = (int)swAnnotationVisibilityState_e.swAnnotationHidden;
                return false;
            }

            double val = tolStorage.GetValue(Unit.Millimeter);

            if (val == 0.0 || double.IsNaN(val))
            {
                ann.Visible = (int)swAnnotationVisibilityState_e.swAnnotationHidden;
                Logger.Info("Datum annotation hidden due to zero tolerance.");
            }
            else
            {
                datumTag.SetLabel(label);

                if (drawDimensions.TryGet("DatumFeature", out var dimAnn) && dimAnn.Position != null)
                {
                    var coords = dimAnn.Position.GetValues(Unit.Meter);
                    ann.SetPosition2(coords[0], coords[1], 0.0);
                    ann.Visible = (int)swAnnotationVisibilityState_e.swAnnotationVisible;
                    Logger.Success("Datum feature label placed.");
                }
                else
                {
                    Logger.Warn("DatumFeature position not found.");
                    return false;
                }
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
            if (_swView?.GetFirstGTOL() is not Gtol gtol)
            {
                Logger.Warn("No GTOL found in view.");
                return false;
            }

            if (gtol.GetAnnotation() is not Annotation ann)
            {
                Logger.Warn("GTOL annotation is null.");
                return false;
            }

            if (!wedgeDimensions.TryGet("SymmetryTolerance", out var tolStorage))
            {
                Logger.Warn("SymmetryTolerance not found.");
                ann.Visible = (int)swAnnotationVisibilityState_e.swAnnotationHidden;
                return false;
            }

            double symTol = tolStorage.GetValue(Unit.Millimeter);

            if (symTol == 0.0 || double.IsNaN(symTol))
            {
                ann.Visible = (int)swAnnotationVisibilityState_e.swAnnotationHidden;
                Logger.Info("GTOL annotation hidden due to zero tolerance.");
                return true;
            }

            if (!drawDimensions.TryGet("GeometricTolerance", out var dimAnn) || dimAnn.Position == null)
            {
                Logger.Warn("GeometricTolerance position not found.");
                return false;
            }

            var coords = dimAnn.Position.GetValues(Unit.Meter);
            gtol.SetPosition(coords[0], coords[1], 0.0);

            string tolInInch = Math.Round(symTol, 4).ToString("0.0000");
            string tolInMm = "[" + symTol.ToString("0.###") + "]";

            bool result = gtol.SetFrameValues2(1, tolInInch, "", tolInMm, label, "");

            if (result)
                Logger.Success("GTOL frame set successfully.");
            else
                Logger.Warn("GTOL frame setup failed.");

            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to place GTOL frame: {ex.Message}");
            return false;
        }
    }

    public bool CreateDatumFeatureSymbol(string label, double x, double y)
    {
        try
        {
            if (_swView == null || _model == null)
            {
                Logger.Warn("View or model is null.");
                return false;
            }

            // Select the view before inserting datum tag
            bool status = _model.Extension.SelectByID2(
                _swView.Name,
                "DRAWINGVIEW",
                0, 0, 0,
                false,
                0,
                null,
                0
            );

            if (!status)
            {
                Logger.Warn("Failed to select view before inserting datum tag.");
                return false;
            }

            // Cast model as DrawingDoc
            var drawingDoc = _model as DrawingDoc;
            if (drawingDoc == null)
            {
                Logger.Error("Model is not a DrawingDoc.");
                return false;
            }

            // Insert datum tag (void return)
            drawingDoc.InsertDatumTag();

            // Get newly inserted datum tag
            DatumTag datumTag = (DatumTag)_swView.GetFirstDatumTag();
            if (datumTag == null)
            {
                Logger.Error("Failed to retrieve newly inserted datum tag.");
                return false;
            }

            // Set label
            datumTag.SetLabel(label);

            // Get annotation
            var ann = datumTag.GetAnnotation() as Annotation;
            if (ann == null)
            {
                Logger.Error("Failed to get annotation from datum tag.");
                return false;
            }

            // Set position (in meters)
            ann.SetPosition2(x, y, 0.0);
            ann.Visible = (int)swAnnotationVisibilityState_e.swAnnotationVisible;

            Logger.Success($"Datum feature symbol '{label}' placed at ({x:F4}, {y:F4}).");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to create datum feature symbol: {ex.Message}");
            return false;
        }
    }

}
