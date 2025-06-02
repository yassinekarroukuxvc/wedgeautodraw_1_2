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

public class BreaklineHandler
{
    private readonly View _swView;
    private readonly ModelDoc2 _model;

    public BreaklineHandler(View swView, ModelDoc2 model)
    {
        _swView = swView;
        _model = model;
    }

    public bool SetBreaklineGap(double gap)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot set breakline gap. View is null.");
            return false;
        }

        try
        {
            _swView.BreakLineGap = gap;
            Logger.Info($"Set breakline gap to {gap:F3} meters.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to set breakline gap: {ex.Message}");
            return false;
        }
    }

    public bool SetBreaklinePosition(NamedDimensionValues wedgeDimensions, DrawingData drawData)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot set breakline position. View is null.");
            return false;
        }

        try
        {
            double scale = _swView.ScaleDecimal;
            double tl = wedgeDimensions["TL"].GetValue(Unit.Meter);
            BreakLine breakLine = _swView.IGetBreakLines(_swView.GetBreakLineCount2(out _));

            if (breakLine == null)
            {
                Logger.Warn("No breakline object found.");
                return false;
            }

            string viewName = _swView.Name.ToLower();
            (double lower, double upper, bool isDetail) = viewName switch
            {
                "front_view" => (drawData.BreaklineData["Front_viewLowerPartLength"].GetValue(Unit.Meter), drawData.BreaklineData["Front_viewLowerPartLength"].GetValue(Unit.Meter), false),
                "side_view" => (drawData.BreaklineData["Side_viewLowerPartLength"].GetValue(Unit.Meter), drawData.BreaklineData["Side_viewLowerPartLength"].GetValue(Unit.Meter), false),
                "detail_view" => (drawData.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Meter), tl / 2, true),
                _ => (0, 0, false)
            };

            double[] breaklinePos = isDetail
                ? new[] { lower - scale * tl / 2, scale * upper }
                : new[] { -tl * scale / 2 + lower, tl * scale / 2 - upper };

            bool result = breakLine.SetPosition(breaklinePos[0], breaklinePos[1]);

            if (isDetail)
            {
                _model.Extension.SelectByID2("TL@Detail_View", "DIMENSION", 0, 0, 0, false, 0, null, 0);
                bool deletion = _model.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Advanced);
                result &= deletion;
            }

            Logger.Info($"Breakline position set successfully in {viewName} view.");
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error setting breakline position: {ex.Message}");
            return false;
        }
    }
    public bool SetOverlayBreaklineRightShift(double shiftAmount = 0.005)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot shift breakline. View is null.");
            return false;
        }

        try
        {
            int count = _swView.GetBreakLineCount2(out _);
            if (count == 0)
            {
                Logger.Warn("No breakline exists in the view to adjust.");
                return false;
            }

            BreakLine breakLine = _swView.IGetBreakLines(count);
            if (breakLine == null)
            {
                Logger.Warn("Breakline object could not be retrieved.");
                return false;
            }

            // Get original positions for both ends of the breakline
            double lowerPos = breakLine.GetPosition(0); 
            double upperPos = breakLine.GetPosition(1);

            // Shift both sides to the right
            double newLowerPos = lowerPos + shiftAmount;
            double newUpperPos = upperPos + shiftAmount;

            bool success = breakLine.SetPosition(newLowerPos, newUpperPos);

            if (success)
            {
                Logger.Info($"Overlay breakline shifted by {shiftAmount} meters. New positions: ({newLowerPos:F4}, {newUpperPos:F4})");
            }
            else
            {
                Logger.Warn("Breakline position update failed.");
            }

            return success;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error shifting overlay breakline: {ex.Message}");
            return false;
        }
    }

    public bool SetOverlayBreaklinePosition(NamedDimensionValues wedgeDimensions, DrawingData drawData)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot set breakline position. View is null.");
            return false;
        }

        try
        {
            string viewName = _swView.Name;


            // Applies only to overlay views
            if (viewName != Constants.OverlayDetailView && viewName != Constants.OverlaySectionView)
            {
                Logger.Warn($"SetOverlayBreaklinePosition skipped: {viewName} is not an overlay view.");
                return false;
            }

            double scale = _swView.ScaleDecimal;
            double tl = wedgeDimensions["TL"].GetValue(Unit.Meter);
            BreakLine breakLine = _swView.IGetBreakLines(_swView.GetBreakLineCount2(out _));

            // Load lower part reference
            string breakKey = viewName == Constants.OverlayDetailView
                ? "Detail_viewLowerPartLength"
                : "Section_viewLowerPartLength";

            if (!drawData.BreaklineData.TryGet(breakKey, out var lowerStorage))
            {
                Logger.Warn($"{breakKey} not found.");
                return false;
            }


            /////

            var x = drawData.BreaklineData["Detail_viewLowerPartLength"].GetValue(Unit.Meter);
            var y = tl / 2;
            bool isDetail = true;
            double[] breaklinePos = isDetail
                ? new[] { x - scale * tl * 30 / 2, scale * y * 100 }
                : new[] { -tl * scale / 2 + y * 100, tl * scale / 2 - y };

            ////




            double lower = lowerStorage.GetValue(Unit.Meter);
            double visibleLength = 0.01; // Show only 10mm of detail in drawing
            double upper = lower + visibleLength;

            // Shift both to drawing space using view scale
            double lowerPos = scale * lower;
            double upperPos = scale * upper;


            if (breakLine == null)
            {
                Logger.Warn("No breakline object found in overlay view.");
                return false;
            }

            bool result = breakLine.SetPosition(breaklinePos[0], breaklinePos[1]);

            Logger.Info($"Overlay breakline set in {viewName} to show range: {breaklinePos[0]:F4} m to {breaklinePos[1]:F4} m");
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error setting overlay breakline position: {ex.Message}");
            return false;
        }
    }



}