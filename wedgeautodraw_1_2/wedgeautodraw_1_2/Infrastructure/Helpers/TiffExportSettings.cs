using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public class TiffExportSettings
{
    private readonly SldWorks _swApp;
    private readonly ModelDoc2 _model;
    public TiffExportSettings(SldWorks swApp)
    {
        _swApp = swApp;
    }
    public bool SetTiffExportSettings(int dpi = 100)
    {
        // 640×480 pixels at 100 DPI → 0.16256 m × 0.12192 m
        double widthMeters = 640 / 100.0 * 0.0254;  // = 0.16256
        double heightMeters = 480 / 100.0 * 0.0254; // = 0.12192


        try
        {
            Logger.Info($"Setting TIFF export settings: DPI={dpi}, Width={widthMeters:F5}m, Height={heightMeters:F5}m");
            _swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture, 1);

            bool dpiSet = _swApp.SetUserPreferenceIntegerValue(
                (int)swUserPreferenceIntegerValue_e.swTiffPrintDPI, dpi);

            bool widthSet = _swApp.SetUserPreferenceDoubleValue(
                (int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperWidth, widthMeters);

            bool heightSet = _swApp.SetUserPreferenceDoubleValue(
                (int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperHeight, heightMeters);

            if (dpiSet && widthSet && heightSet)
            {
                Logger.Success("TIFF export settings applied successfully.");
                return true;
            }

            Logger.Warn("Some TIFF export preferences failed to apply.");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to set TIFF export preferences: {ex.Message}");
            return false;
        }
    }

    public void PrintTiffExportSettings()
    {
        try
        {
            int dpi = _swApp.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffPrintDPI);
            double width = _swApp.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperWidth);
            double height = _swApp.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperHeight);

            Logger.Info($"Current TIFF Export Settings:");
            Logger.Info($"  DPI           : {dpi}");
            Logger.Info($"  Paper Width   : {width:F4} meters ({width * 39.3701:F2} inches)");
            Logger.Info($"  Paper Height  : {height:F4} meters ({height * 39.3701:F2} inches)");

            double widthPx = width * dpi / 0.0254;
            double heightPx = height * dpi / 0.0254;

            Logger.Info($"  Pixel Size    : {Math.Round(widthPx)} x {Math.Round(heightPx)} pixels");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to read TIFF export preferences: {ex.Message}");
        }
    }

    public bool RunSolidWorksMacro(SldWorks swApp, string macroPath)
    {
        int errorCode;
        bool result = swApp.RunMacro2(
            macroPath,
            "Macro1", // Your module name in the .swp macro
            "main", // Your subroutine name in the macro
            (int)swRunMacroOption_e.swRunMacroUnloadAfterRun,
            out errorCode
        );

        if (result)
            Logger.Success("Macro executed successfully.");
        else
            Logger.Error($"Macro execution failed with error code: {errorCode}");

        return result;
    }

}
