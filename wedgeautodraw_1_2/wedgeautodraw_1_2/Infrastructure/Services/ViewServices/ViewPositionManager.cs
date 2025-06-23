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
    public bool MoveViewToPosition(double targetX_inch, double targetY_inch)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot move view. View is null.");
            return false;
        }

        try
        {
            // Convert inches to meters
            const double inchToMeter = 0.0254;
            double targetX = targetX_inch * inchToMeter;
            double targetY = targetY_inch * inchToMeter;

            object posObj = _swView.Position;
            double[] currentPos = posObj as double[];

            if (currentPos == null || currentPos.Length < 2)
            {
                Logger.Warn("Failed to get valid view position.");
                return false;
            }

            Logger.Info($"Current view position: X = {currentPos[0]:F4} m, Y = {currentPos[1]:F4} m");

            _swView.Position = new double[] { targetX, targetY };

            Logger.Success($"Moved view '{_swView.Name}' to: X = {targetX:F4} m, Y = {targetY:F4} m");

            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception during MoveViewToPosition: {ex.Message}");
            return false;
        }
    }
    public bool ShiftViewRight(double shiftAmount)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot shift view. View is null.");
            return false;
        }

        try
        {
            shiftAmount = shiftAmount / 1000;
            // Get current position of the view (in meters)
            double[] currentPos = (double[])_swView.Position;

            double currentX = currentPos[0];
            double currentY = currentPos[1];

            double newX = currentX + shiftAmount;
            double newY = currentY; // Keep Y unchanged

            // Apply new position
            _swView.Position = new double[] { newX, newY };

            Logger.Info($"Shifted view '{_swView.Name}' to the right by {shiftAmount:F4} meters. New position: ({newX:F4}, {newY:F4})");

            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error while shifting view: {ex.Message}");
            return false;
        }
    }
    public bool CenterViewVerticallyByFL(NamedDimensionValues wedgeDimensions)
    {
        if (_swView == null || _drawingDoc == null)
        {
            Logger.Warn("View or drawing is null. Cannot center.");
            return false;
        }

        try
        {
            // 1. Get current sheet height
            SolidWorks.Interop.sldworks.Sheet sheet = (SolidWorks.Interop.sldworks.Sheet)_drawingDoc.GetCurrentSheet();
            double sheetWidthInches = 0, sheetHeightInches = 0;
            sheet.GetSize(ref sheetWidthInches, ref sheetHeightInches);
            double sheetCenterY_m = (sheetHeightInches * 0.0254) / 2.0;

            // 2. Get the current bounding box of the view geometry
            double[] outline = (double[])_swView.GetOutline();
            double geometryCenterY = (outline[1] + outline[3]) / 2.0;

            // 3. Get current view position
            double[] currentPos = (double[])_swView.Position;
            double currentViewY = currentPos[1];

            // 4. Calculate shift needed to align geometry center with sheet center
            double shiftY = sheetCenterY_m - geometryCenterY;
            double newY = currentViewY + shiftY;

            // 5. Apply the shift
            _swView.Position = new double[] { 100, newY };

            Logger.Success($"View '{_swView.Name}' centered vertically using geometry center Y = {geometryCenterY:F4}, sheet Y center = {sheetCenterY_m:F4}.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to center view using geometry: {ex.Message}");
            return false;
        }
    }
    public DataStorage GetVerticalSheetCenterPosition(DrawingData drawData)
    {
        if (_drawingDoc == null)
        {
            Logger.Warn("Cannot calculate sheet center. Drawing document is null.");
            return null;
        }

        try
        {
            // 1. Get sheet size in meters
            var currentSheet = (SolidWorks.Interop.sldworks.Sheet)_drawingDoc.GetCurrentSheet();
            double sheetWidth_m = 0, sheetHeight_m = 0;
            currentSheet.GetSize(ref sheetWidth_m, ref sheetHeight_m);
            Logger.Warn($"Sheet size: width = {sheetWidth_m:F4} m, height = {sheetHeight_m:F4} m");

            // 2. Compute center Y in millimeters
            double centerY_mm = (sheetHeight_m / 2.0) * 1000.0;

            // 3. Get X from Detail_view and convert to mm
            double[] xy_m = drawData.ViewPositions["Detail_view"].GetValues(Unit.Meter);
            if (xy_m.Length < 1)
            {
                Logger.Warn("Detail_view position is invalid or empty.");
                return null;
            }

            double x_mm = xy_m[0] * 1000.0;

            // 4. Return new DataStorage in millimeters
            return new DataStorage(new[] { x_mm, centerY_mm });
        }
        catch (Exception ex)
        {
            Logger.Error($"Error calculating vertical center position: {ex.Message}");
            return null;
        }
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
    public void CenterViewHorizontally(bool isDetailView)
    {
        if (_drawingDoc == null || _swView == null)
        {
            Logger.Warn("Cannot center view horizontally. View or drawing is null.");
            return;
        }

        try
        {
            // 1. Get sheet size
            var sheet = (Sheet)_drawingDoc.GetCurrentSheet();
            double sheetWidth = 0, sheetHeight = 0;
            sheet.GetSize(ref sheetWidth, ref sheetHeight);
            double targetCenterX = isDetailView ? sheetWidth- 0.0127 : sheetWidth / 2.0;

            // 2. Compute visible model length
            const double safetyMargin = 0.065;
            double scale = _swView.ScaleDecimal;
            double visibleLength_m = (sheetWidth / 2.0 - safetyMargin) / scale;
            visibleLength_m += 0.000025;
            double shiftLeft = visibleLength_m / 2.0 * scale;

            // 3. Get view bounding box and current position
            double[] box = (double[])_swView.GetOutline(); // [left, bottom, right, top]
            double viewCenterX = (box[0] + box[2]) / 2.0;
            double[] currentPos = (double[])_swView.Position;

            // 4. Align view center to targetCenterX (center or right edge)
            double shiftToTargetCenter = targetCenterX - viewCenterX;
            double[] centeredPos = new[] { currentPos[0] + shiftToTargetCenter, currentPos[1] };
            _swView.Position = centeredPos;

            // 5. Shift left by half the visible model length
            double[] finalPos = new[] { centeredPos[0] - shiftLeft, centeredPos[1] };
            _swView.Position = finalPos;

            Logger.Success($"{(isDetailView ? "Detail" : "Section")} view aligned to {(isDetailView ? "right edge" : "center")} then shifted left by {shiftLeft * 1000:F2} mm.");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to center and shift view: {ex.Message}");
        }
    }

    public void CenterViewHorizontally2()
    {
        if (_drawingDoc == null || _swView == null)
        {
            Logger.Warn("Cannot center view horizontally. View or drawing is null.");
            return;
        }

        try
        {
            // 1. Get sheet size
            var sheet = (Sheet)_drawingDoc.GetCurrentSheet();
            double sheetWidth = 0, sheetHeight = 0;
            sheet.GetSize(ref sheetWidth, ref sheetHeight);
            double sheetCenterX = sheetWidth / 2.0;

            // 2. Get view outline (in meters)
            double[] box = (double[])_swView.GetOutline(); // [left, bottom, right, top]
            double viewCenterX = (box[0] + box[2]) / 2.0;
            double viewRightX = box[2];

            // 3. Get current position
            double[] pos = (double[])_swView.Position;

            // 4. Center the view
            double shiftToCenter = sheetCenterX - viewCenterX;
            double[] newPos = new[] { pos[0] + shiftToCenter, pos[1] };
            _swView.Position = newPos;

            // 5. Compute new view right edge after centering
            box = (double[])_swView.GetOutline(); // [left, bottom, right, top]
            viewCenterX = (box[0] + box[2]) / 2.0;
            viewRightX = box[2];
            double correctedViewRightX = viewRightX;

            // 6. Shift left by (viewRightX - sheetCenterX)
            double shiftLeft = correctedViewRightX - sheetCenterX;
            _swView.Position = new[] { newPos[0] + viewCenterX, newPos[1] };

            Logger.Success($"View '{_swView.Name}' centered, then shifted left by {shiftLeft * 1000:F2} mm so its RIGHT edge aligns with sheet center.");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to center and shift view horizontally: {ex.Message}");
        }
    }

    public void AlignViewRightEdgeToTarget(bool alignToSheetCenter)
    {
        if (_drawingDoc == null || _swView == null)
        {
            Logger.Warn("Cannot align view. Drawing or view is null.");
            return;
        }

        try
        {
            // Get sheet size
            SolidWorks.Interop.sldworks.Sheet sheet = (SolidWorks.Interop.sldworks.Sheet)_drawingDoc.GetCurrentSheet();
            double sheetWidth = 0, sheetHeight = 0;
            sheet.GetSize(ref sheetWidth, ref sheetHeight);

            // Target position: either center or full width
            double targetX = alignToSheetCenter ? (sheetWidth / 2.0) : sheetWidth;

            // Get view bounding box
            double[] box = (double[])_swView.GetOutline();
            double viewRightX = box[2]; // right edge of the view

            // Compute shift
            double[] currentPos = (double[])_swView.Position;
            double shiftX = targetX - viewRightX;

            // Apply shift
            _swView.Position = new[] { currentPos[0] + shiftX, currentPos[1] };

            Logger.Success($"Aligned view's RIGHT edge to {(alignToSheetCenter ? "sheet center" : "sheet right edge")} (shift = {shiftX * 1000:F2} mm)");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to align view: {ex.Message}");
        }
    }
    public void AlignViewRightEdgeToTargetX(double targetX)
    {
        if (_drawingDoc == null || _swView == null)
        {
            Logger.Warn("Cannot align view. Drawing or view is null.");
            return;
        }

        try
        {
            // Get current bounding box
            double[] box = (double[])_swView.GetOutline();
            double viewRightX = box[2];

            // Compute horizontal shift
            double[] pos = (double[])_swView.Position;
            double shiftX = targetX - viewRightX;

            _swView.Position = new[] { pos[0] + shiftX, pos[1] };

            Logger.Success($"Aligned view '{_swView.Name}' right edge to X = {targetX:F4} m (shift = {shiftX * 1000:F2} mm)");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to align view right edge to target: {ex.Message}");
        }
    }
    public bool SetViewX(double x_mm)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot set view X position. View is null.");
            return false;
        }

        try
        {
            double x_m = x_mm / 1000.0; // Convert mm → meters
            double[] currentPos = (double[])_swView.Position;

            _swView.Position = new[] { x_m, currentPos[1] };

            Logger.Success($"Set view '{_swView.Name}' X position to {x_mm:F2} mm ({x_m:F4} m)");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to set view X position: {ex.Message}");
            return false;
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
