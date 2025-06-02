using DocumentFormat.OpenXml.EMMA;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace wedgeautodraw_1_2.Infrastructure.Services;

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
    public void SaveAsTiff(string outputPath)
    {
        try
        {
            bool success = _swModelExt.SaveAs(
                outputPath,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                null,
                ref _error,
                ref _warning
            );

            if (success)
                Logger.Success($"Drawing saved as TIF: {outputPath}");
            else
                Logger.Error($"Failed to save TIF. Error code: {_error}, Warning code: {_warning}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception during TIF export: {ex.Message}");
        }
    }
    public void SaveAsPdfAndConvertToTiff(string pdfPath, string tiffPath, int dpi = 300)
    {
        try
        {
            // Step 1: Save as PDF
            Logger.Info($"Saving drawing as PDF: {pdfPath}");
            SaveAsPdf(pdfPath);

            // Step 2: Convert PDF → TIFF using Magick.NET
            Logger.Info($"Converting PDF to TIFF at {dpi} DPI...");

            using (var images = new ImageMagick.MagickImageCollection())
            {
                // Set read settings
                var settings = new ImageMagick.MagickReadSettings
                {
                    Density = new ImageMagick.Density(dpi) // Controls DPI of the output TIFF
                };

                // Read PDF pages
                images.Read(pdfPath, settings);

                // Optionally, flatten multi-page PDF to one TIFF page
                using (var merged = images.AppendVertically())
                {
                    merged.Format = ImageMagick.MagickFormat.Tiff;
                    merged.Write(tiffPath);
                }
            }

            Logger.Success($"PDF converted to TIFF successfully: {tiffPath}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception during PDF → TIFF conversion: {ex.Message}");
        }
    }


}
