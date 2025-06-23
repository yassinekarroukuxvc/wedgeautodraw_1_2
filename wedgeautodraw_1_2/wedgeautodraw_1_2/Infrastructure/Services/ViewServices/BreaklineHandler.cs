using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
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
            var breakLine = GetBreaklineObject();
            if (breakLine == null) return false;

            string viewName = _swView.Name.ToLower();
            double scale = _swView.ScaleDecimal;
            double tl = wedgeDimensions["TL"].GetValue(Unit.Meter);

            if (!TryGetBreaklineConfig(viewName, drawData, tl, scale, out var pos, out bool isDetail))
                return false;

            bool result = breakLine.SetPosition(pos[0], pos[1]);

            if (isDetail)
            {
                _model.Extension.SelectByID2("TL@Detail_View", "DIMENSION", 0, 0, 0, false, 0, null, 0);
                result &= _model.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Advanced);
            }

            Logger.Info($"Breakline position set successfully in '{viewName}' view.");
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
        var breakLine = GetBreaklineObject();
        if (breakLine == null) return false;

        try
        {
            double lower = breakLine.GetPosition(0);
            double upper = breakLine.GetPosition(1);

            bool success = breakLine.SetPosition(lower + shiftAmount, upper + shiftAmount);

            Logger.Info(success
                ? $"Overlay breakline shifted by {shiftAmount:F4} m → New: ({lower + shiftAmount:F4}, {upper + shiftAmount:F4})"
                : "Failed to shift overlay breakline.");

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
        var breakLine = GetBreaklineObject();
        if (breakLine == null) return false;

        try
        {
            if (_model is not DrawingDoc drawingDoc)
            {
                Logger.Warn("Model is not a DrawingDoc.");
                return false;
            }

            double scale = _swView.ScaleDecimal;
            double tl = wedgeDimensions["TL"].GetValue(Unit.Meter);

            var sheet = (Sheet)drawingDoc.GetCurrentSheet();
            double sheetWidth = 0, sheetHeight = 0;
            sheet.GetSize(ref sheetWidth, ref sheetHeight);


            double safetyMargin = 0.065;
            double visibleLength_m = (sheetWidth / 2.0 - safetyMargin) / scale;

            double lower = visibleLength_m * scale;
            double upper = visibleLength_m * scale;

            double[] pos = {
                lower - scale * tl / 2,
                upper - scale * tl / 2
            };

            bool result = breakLine.SetPosition(pos[0], pos[1]);

            Logger.Info(result
                ? $"Overlay breakline positioned → Lower: {pos[0]:F4}, Upper: {pos[1]:F4}"
                : "Failed to position overlay breakline.");

            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error in SetOverlayBreaklinePosition: {ex.Message}");
            return false;
        }
    }

    private BreakLine GetBreaklineObject()
    {
        if (_swView == null)
        {
            Logger.Warn("View is null. Cannot get breakline.");
            return null;
        }

        int count = _swView.GetBreakLineCount2(out _);
        if (count == 0)
        {
            Logger.Warn("No breaklines in view.");
            return null;
        }

        var breakLine = _swView.IGetBreakLines(count);
        if (breakLine == null)
            Logger.Warn("Failed to retrieve breakline object.");

        return breakLine;
    }

    private bool TryGetBreaklineConfig(string viewName, DrawingData drawData, double tl, double scale, out double[] pos, out bool isDetail)
    {
        isDetail = viewName == "detail_view";

        string baseKey = viewName switch
        {
            "front_view" => "Front_view",
            "side_view" => "Side_view",
            "detail_view" => "Detail_view",
            _ => null
        };

        if (baseKey == null || !drawData.BreaklineData.ContainsKey($"{baseKey}LowerPartLength"))
        {
            Logger.Warn($"Breakline config not found for view: {viewName}");
            pos = new[] { 0.0, 0.0 };
            return false;
        }

        double lower = drawData.BreaklineData[$"{baseKey}LowerPartLength"].GetValue(Unit.Meter);
        double upper = isDetail
            ? tl / 2
            : drawData.BreaklineData.TryGet($"{baseKey}UpperPartLength", out var upperVal)
                ? upperVal.GetValue(Unit.Meter)
                : lower;

        pos = isDetail
            ? new[] { lower - scale * tl / 2, scale * upper }
            : new[] { -tl * scale / 2 + lower, tl * scale / 2 - upper };

        return true;
    }
}
