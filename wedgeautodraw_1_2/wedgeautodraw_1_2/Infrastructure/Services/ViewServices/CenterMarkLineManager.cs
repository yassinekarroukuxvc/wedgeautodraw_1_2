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

class CenterMarkLineManager
{
    private readonly View _swView;
    private readonly ModelDoc2 _model;
    public CenterMarkLineManager(View swView, ModelDoc2 model)
    {
        _swView = swView;
        _model = model;
    }
    public bool CreateCenterline(NamedDimensionValues wedgeDimensions, DrawingData drawData)
    {
        try
        {
            string viewName = _swView.Name.ToLower();
            double scale = _swView.ScaleDecimal;
            bool isOverlay = drawData.DrawingType == DrawingType.Overlay;

            double ScaleOffset(double mm) => mm / 1000 / scale;

            double[] GetCenterlineCoordinates()
            {
                switch (viewName)
                {
                    case "front_view":
                    case "side_view":
                        double tl = wedgeDimensions["TL"].GetValue(Unit.Meter);

                        // ↕️ Production: vertical / ↔️ Overlay: horizontal
                        if (!isOverlay)
                        {
                            return new[] { 0.0, tl / 2 + ScaleOffset(2), 0.0, -tl / 2 - ScaleOffset(2) };
                        }
                        else
                        {
                            return new[] { -tl / 2 - ScaleOffset(2), 0.0, tl / 2 + ScaleOffset(2), 0.0 };
                        }

                    case "detail_view":
                        double tlDetail = wedgeDimensions["TL"].GetValue(Unit.Meter);
                        return new[] { 0.0, 0.0, 0.0, -tlDetail / 2 + ScaleOffset(10) };

                    case "drawing view2":
                    case "drawing view1":
                        tl = wedgeDimensions["TL"].GetValue(Unit.Meter);
                        if (!isOverlay)
                        {
                            return new[] { 0.0, tl / 2, 0.0, -tl / 2 };
                        }
                        else
                        {
                            return new[] { -tl / 2 - ScaleOffset(2), 0.0, tl / 2 + ScaleOffset(2), 0.0 };
                        }

                    default:
                        return null;
                }
            }

            double[] pos = GetCenterlineCoordinates();
            if (pos == null) return false;

            SketchSegment line = _model.SketchManager.CreateCenterLine(
                pos[0], pos[1], 0.0,
                pos[2], pos[3], 0.0);

            line.Layer = "center";
            line.GetSketch().RelationManager.AddRelation(
                new[] { line }, (int)swConstraintType_e.swConstraintType_FIXED);

            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"CreateCenterline failed: {ex.Message}");
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
                line.Layer = "center";
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

    public bool CreateCenterlineAtViewCenter(bool isVertical = true, double lengthMm = 100.0)
    {
        if (_swView == null || _model == null)
        {
            Logger.Warn("Model or View is null. Cannot draw centerline.");
            return false;
        }

        try
        {
            double scale = _swView.ScaleDecimal;
            double ScaleOffset(double mm) => mm / 1000.0 / scale;
            double halfLength = ScaleOffset(lengthMm) / 2.0;

            // Select the drawing view
            bool selected = _model.Extension.SelectByID2(
                _swView.Name,
                "DRAWINGVIEW",
                0, 0, 0, false, 0, null, 0
            );

            if (!selected)
            {
                Logger.Warn($"Could not select view {_swView.Name}.");
                return false;
            }

            // Enter sketch mode on the view
            ISketchManager sketchMgr = _model.SketchManager;
            sketchMgr.InsertSketch(true); // open sketch

            // Draw centerline at (0,0) in view coordinates
            double x1, y1, x2, y2;

            if (isVertical)
            {
                x1 = 0;
                y1 = -halfLength;
                x2 = 0;
                y2 = halfLength;
            }
            else
            {
                x1 = -halfLength;
                y1 = 0;
                x2 = halfLength;
                y2 = 0;
            }

            SketchSegment line = sketchMgr.CreateCenterLine(x1, y1, 0.0, x2, y2, 0.0);

            sketchMgr.InsertSketch(true); // exit sketch

            if (line != null)
            {
                line.Layer = "center";
                line.GetSketch().RelationManager.AddRelation(new[] { line }, (int)swConstraintType_e.swConstraintType_FIXED);
                Logger.Success($"Centerline successfully created in view sketch of '{_swView.Name}'.");
                return true;
            }

            Logger.Warn("Centerline creation failed inside view sketch.");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Error($"CreateCenterlineInViewSketchSpace failed: {ex.Message}");
            return false;
        }
    }

}
