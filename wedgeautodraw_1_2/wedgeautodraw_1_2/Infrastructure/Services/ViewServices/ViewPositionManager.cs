using DocumentFormat.OpenXml.EMMA;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services.ViewServices;

public class ViewPositionManager
{
    private readonly View _swView;
    private readonly ModelDoc2 _model;
    private readonly DrawingDoc _drawingDoc;
    public ViewPositionManager(View swView, ModelDoc2 model, DrawingDoc drawingDoc)
    {
        _swView = swView;
        _model = model;
        _drawingDoc = drawingDoc;
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
    public double[] GetPosition()
    {
        if (_swView?.Position is double[] position && position.Length >= 2)
        {
            Logger.Info($"Section View Position: X = {position[0]:F4} m, Y = {position[1]:F4} m");
            return position;
        }

        return new double[] { 0.0, 0.0 };
    }
    public void CenterViewVertically()
    {
        if (_drawingDoc == null || _swView == null)
        {
            Logger.Warn("Cannot center view vertically. View or drawing is null.");
            return;
        }

        try
        {
            SolidWorks.Interop.sldworks.Sheet sheet = (SolidWorks.Interop.sldworks.Sheet)_drawingDoc.GetCurrentSheet();
            double sheetWidth = 0, sheetHeight = 0;
            sheet.GetSize(ref sheetWidth, ref sheetHeight);

            double sheetCenterY = sheetHeight / 2.0;

            double[] box = (double[])_swView.GetOutline();
            double viewBottomY = box[1];
            double viewTopY = box[3];
            double viewCenterY = (viewBottomY + viewTopY) / 2.0;

            double[] pos = (double[])_swView.Position;
            double shiftY = sheetCenterY - viewCenterY;
            _swView.Position = new[] { pos[0], pos[1] + shiftY };

            Logger.Success($"Vertically centered view (Y shift = {shiftY * 1000:F2} mm)");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to center view vertically: {ex.Message}");
        }
    }
    public void CenterViewHorizontally(bool isDetailView, double tlInMeters = 0)
    {
        if (_drawingDoc == null || _swView == null)
        {
            Logger.Warn("Cannot center view horizontally. View or drawing is null.");
            return;
        }

        try
        {
            Sheet sheet = (Sheet)_drawingDoc.GetCurrentSheet();
            double sheetWidth = 0, sheetHeight = 0;
            sheet.GetSize(ref sheetWidth, ref sheetHeight);
            double scale = _swView.ScaleDecimal;

            if (tlInMeters != 0)
            {
                double tlScaled = tlInMeters * scale;

                // Get current view bounding box and position
                double[] box = (double[])_swView.GetOutline(); // [left, bottom, right, top]
                double viewCenterX = (box[0] + box[2]) / 2.0;
                double viewWidth = (box[2] - box[0]);
                double[] currentPos = (double[])_swView.Position;

                // Compute current left edge in model space
                double currentLeft = viewCenterX - (tlScaled / 2.0);

                // Compute target left edge (we want to align it with sheet center)
                double targetLeft = sheetWidth / 2.0;

                // Determine how much to shift the view
                double shift = targetLeft - currentLeft;

                // Apply new position
                double[] newPos = new[] { currentPos[0] + shift+0.01, currentPos[1] };
                _swView.Position = newPos;

                Logger.Success($"Side view aligned: left edge aligned with sheet center (shifted by {shift * 1000:F2} mm).");
                Logger.Error("DEEZ NUts");
            }
            else
            {
                // --- DETAIL / SECTION VIEW LOGIC ---
                double targetCenterX = isDetailView ? sheetWidth - 0.005 : sheetWidth / 2.0;

                const double safetyMargin = 0.140;
                double visibleLength_m = (sheetWidth / 2.0 - safetyMargin) / scale;
                visibleLength_m += 0.000025;
                double shiftLeft = visibleLength_m / 2.0 * scale;

                double[] box = (double[])_swView.GetOutline(); // [left, bottom, right, top]
                double viewCenterX = (box[0] + box[2]) / 2.0;
                double[] currentPos = (double[])_swView.Position;

                double shiftToTargetCenter = targetCenterX - viewCenterX;
                double[] centeredPos = new[] { currentPos[0] + shiftToTargetCenter, currentPos[1] };
                _swView.Position = centeredPos;

                double[] finalPos = new[] { centeredPos[0] - shiftLeft, centeredPos[1] };
                _swView.Position = finalPos;

                Logger.Success($"{(isDetailView ? "Detail" : "Section")} view centered then shifted left by {shiftLeft * 1000:F2} mm.");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to center view: {ex.Message}");
        }
    }
    public bool CenterSectionViewVisuallyVertically(NamedDimensionValues wedgeDimensions)
    {
        if (_swView == null || _drawingDoc == null)
        {
            Logger.Warn("Cannot vertically correct section view. View or drawing is null.");
            return false;
        }

        try
        {
            // 1. Get current position
            double[] pos = (double[])_swView.Position;
            double currentX = pos[0];
            double currentY = pos[1];

            // 2. Get TDF and define side edge offset (in mm → convert to meters)
            double tdf = wedgeDimensions["TDF"].GetValue(Unit.Meter);
            double td = wedgeDimensions["TD"].GetValue(Unit.Meter);
            const double sideEdgeOffset_mm = 6.0;
            double sideEdgeOffset_m = sideEdgeOffset_mm / 1000.0;

            // 3. Correction formula (user-defined based on visual test)
            double correctionY = (td - tdf) / 2 * _swView.ScaleDecimal;

            // 4. Apply corrected Y
            double newY = currentY - correctionY;

            _swView.Position = new[] { currentX, newY };

            Logger.Success($"Section view Y visually centered with geometry correction. ΔY = {correctionY * 1000:F2} mm");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error while visually centering section view in Y: {ex.Message}");
            return false;
        }
    }
}
