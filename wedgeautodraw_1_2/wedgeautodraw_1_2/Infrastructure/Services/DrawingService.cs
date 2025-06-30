using DocumentFormat.OpenXml.EMMA;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;
namespace wedgeautodraw_1_2.Infrastructure.Services;
using System.Runtime.InteropServices;

public class DrawingService : IDrawingService
{
    private readonly SldWorks _swApp;
    private ModelDoc2 _swModel;
    private DrawingDoc _swDrawing;
    private ModelDocExtension _swModelExt;
    private CustomPropertyManager _swCustProps;
    private string _drawingPath;

    private int _error = 0;
    private int _warning = 0;

    public DrawingService(SldWorks swApp)
    {
        _swApp = swApp;
    }

    public void OpenDrawing(string filePath)
    {
        _drawingPath = filePath;
        _swApp.OpenDoc6(filePath,
                        (int)swDocumentTypes_e.swDocDRAWING,
                        (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                        "",
                        ref _error,
                        ref _warning);

        _swModel = (ModelDoc2)_swApp.ActiveDoc;

        if (_swModel == null)
        {
            Logger.Error($"Failed to open document at: {filePath}");
            throw new InvalidOperationException("No active document after opening.");
        }

        if (_swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
        {
            Logger.Error($"Opened document is not a drawing: {filePath}");
            throw new InvalidCastException("Active document is not a drawing.");
        }

        _swDrawing = _swModel as DrawingDoc;

        if (_swDrawing == null)
        {
            Logger.Error("Casting to DrawingDoc failed.");
            throw new InvalidCastException("Failed to cast ModelDoc2 to DrawingDoc.");
        }
        
        _swModelExt = _swModel.Extension;
        _swModel.Lock();
    }
    public void SaveDrawing()
    {
        _swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref _error, ref _warning);
    }
    public void SaveAndCloseDrawing()
    {
        SaveDrawing();
        _swApp.CloseDoc(_drawingPath);
    }
    public void SaveAsPdf(string outputPath)
    {
        var pdfData = (ExportPdfData)_swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);

        var sheetNames = (string[])_swDrawing.GetSheetNames();
        if (sheetNames == null || sheetNames.Length == 0)
        {
            Logger.Warn("No sheets found to export.");
            return;
        }

        var wrappers = new DispatchWrapper[sheetNames.Length];

        for (int i = 0; i < sheetNames.Length; i++)
        {
            _swDrawing.ActivateSheet(sheetNames[i]);
            var sheet = (Sheet)_swDrawing.GetCurrentSheet();
            wrappers[i] = new DispatchWrapper(sheet);
        }

        pdfData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, wrappers);
        pdfData.ViewPdfAfterSaving = false;

        _swModelExt.SaveAs3(
            outputPath,
            (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
            (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
            pdfData,
            null,
            ref _error,
            ref _warning
        );
    }
    public void SetSummaryInformation(DrawingData drawingData)
    {
        string title = drawingData.TitleInfo["number"] + " - " + drawingData.DrawingType + " Copy";
        string subject = drawingData.DrawingType + " " + drawingData.TitleInfo["number"];

        _swModel.SummaryInfo[(int)swSummInfoField_e.swSumInfoTitle] = title;
        _swModel.SummaryInfo[(int)swSummInfoField_e.swSumInfoSubject] = subject;
        _swModel.SummaryInfo[(int)swSummInfoField_e.swSumInfoSavedBy] = "Autodraw System";
        _swModel.SummaryInfo[(int)swSummInfoField_e.swSumInfoKeywords] = title;
        _swModel.SummaryInfo[(int)swSummInfoField_e.swSumInfoAuthor] = "Autodraw System";
        _swModel.SummaryInfo[(int)swSummInfoField_e.swSumInfoComment] = title + " created by Autodraw Service";
    }
    public void SetCustomProperties(DrawingData drawingData)
    {
        _swCustProps = _swModel.Extension.get_CustomPropertyManager("");

        foreach (var kvp in drawingData.TitleBlockInfo)
        {
            _swCustProps.Set2(kvp.Key, kvp.Value);
        }
    }
    public void Rebuild()
    {
        _swModel.EditRebuild3();
    }
    public void ZoomToFit()
    {
        _swModel.ViewZoomtofit2();
    }
    public void ZoomToSheet()
    {
        try
        {
            if (_swModelExt == null)
            {
                Logger.Warn("Cannot ZoomToSheet: Model extension is null.");
                return;
            }

            _swModelExt.ViewZoomToSheet();
            Logger.Success("Zoomed to sheet (using IModelDocExtension.ViewZoomToSheet).");
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception during ZoomToSheet: {ex.Message}");
        }
    }
    public ModelDoc2 GetModel() => _swModel;
    public void ReplaceReferencedModel(string drawingPath, string oldModelPath, string newModelPath)
    {
        _swApp.ReplaceReferencedDocument(drawingPath, oldModelPath, newModelPath);
    }
    public void Reopen()
    {
        try
        {
            _swApp.OpenDoc6(_drawingPath, (int)swDocumentTypes_e.swDocDRAWING,
                (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, string.Empty, ref _error, ref _warning);
            _swModel = (ModelDoc2)_swApp.ActiveDoc;
            _swDrawing = (DrawingDoc)_swModel;
            _swModelExt = _swModel.Extension;
            _swModel.Lock();
        }
        catch (Exception)
        {
            Logger.Warn("Unable to Reopen the Drawing");
        }
    }
    public void Unlock()
    {
        try
        {
            _swModel.UnLock();
        }
        catch (Exception)
        {
            Logger.Warn("Unable to Unlock the Drawing");
        }
    }
    public void Lock()
    {
        try
        {
            _swModel.Lock();
        }
        catch (Exception)
        {
            Logger.Warn("Unable to Unlock the Drawing");
        }
    }
    public void SaveAsTiff(string outputPath, int dpi = 72, int widthPx = 640, int heightPx = 480)
{
    try
    {
        // Convert pixel dimensions to meters
        double widthM = (widthPx * 25.4 / dpi) / 1000.0;
        double heightM = (heightPx * 25.4 / dpi) / 1000.0;

        Logger.Info($"Exporting TIFF: {widthPx}x{heightPx} px @ {dpi} DPI → {widthM:F4}m x {heightM:F4}m");

        /*// Set preferences (these mostly affect Print Capture, but may still influence SaveAs)
        _swApp.SetUserPreferenceIntegerValue(
            (int)swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture, 1);

        _swApp.SetUserPreferenceIntegerValue(
            (int)swUserPreferenceIntegerValue_e.swTiffPrintDPI, dpi);

        _swApp.SetUserPreferenceIntegerValue(
            (int)swUserPreferenceIntegerValue_e.swTiffImageType,
            (int)swTiffImageType_e.swTiffImageRGB);

        _swApp.SetUserPreferenceIntegerValue(
            (int)swUserPreferenceIntegerValue_e.swTiffCompressionScheme,
            (int)swTiffCompressionScheme_e.swTiffPackbitsCompression);

        _swApp.SetUserPreferenceIntegerValue(
            (int)swUserPreferenceIntegerValue_e.swTiffPrintPaperSize,
            0); // 0 = User Defined

        _swApp.SetUserPreferenceDoubleValue(
            (int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperWidth, widthM);

        _swApp.SetUserPreferenceDoubleValue(
            (int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperHeight, heightM);

        _swApp.SetUserPreferenceToggle(
            (int)swUserPreferenceToggle_e.swTiffPrintUseSheetSize, false); // Use custom size

        _swApp.SetUserPreferenceToggle(
            (int)swUserPreferenceToggle_e.swTiffPrintPadText, true);*/

        // SaveAs to TIFF
        bool success = _swModelExt.SaveAs(
            outputPath,
            (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
            (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
            null,
            ref _error,
            ref _warning
        );

            if (success)
            Logger.Success($"TIFF saved: {outputPath} [{widthPx}x{heightPx} px @ {dpi} DPI]");
        else
            Logger.Error($"TIFF save failed. Error={_error}, Warning={_warning}");
    }
    catch (Exception ex)
    {
        Logger.Error($"TIFF export failed: {ex.Message}");
    }
}
    public void DrawCenteredSquareOnSheet(double sideLengthInInches)
    {
        if (_swDrawing == null || _swModel == null)
        {
            Logger.Error("Cannot draw square: No drawing is open.");
            return;
        }

        try
        {
            string layerName = "square";
            // Convert side length to meters
            double sideLength = sideLengthInInches * 0.0254;

            // Get current sheet (with cast)
            Sheet sheet = (Sheet)_swDrawing.GetCurrentSheet();

            // Enter Edit Sheet Format mode
            _swDrawing.EditTemplate();

            // Set current layer using DrawingDoc.SetCurrentLayer
            _swDrawing.SetCurrentLayer(layerName);
            Logger.Success($"Current layer set to: {layerName}");

            // Get sheet size
            double sheetWidth = 0.0;
            double sheetHeight = 0.0;
            sheet.GetSize(ref sheetWidth, ref sheetHeight);

            // Compute center of sheet
            double centerX = sheetWidth / 2.0;
            double centerY = sheetHeight / 2.0;
            double halfSide = sideLength / 2.0;

            double x1 = centerX - halfSide;
            double y1 = centerY - halfSide;
            double x2 = centerX + halfSide;
            double y2 = centerY + halfSide;

            Logger.Error($"x1 = {x1} y1 = {y1} x2 = {x2} y2 = {y2}");

            // Create rectangle SKETCH
            SketchManager sketchMgr = _swModel.SketchManager;
            sketchMgr.InsertSketch(true);
            sketchMgr.CreateCornerRectangle(x1, y1, 0, x2, y2, 0);
            sketchMgr.InsertSketch(true);

            // Exit Edit Sheet Format mode
            _swDrawing.EditSheet();

            Logger.Success($"Centered square drawn in Sheet Format on layer '{layerName}': {sideLengthInInches} inches side length.");
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception while drawing centered square: {ex.Message}");
        }
    }

}