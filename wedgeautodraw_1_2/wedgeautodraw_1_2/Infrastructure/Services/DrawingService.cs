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
        _swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref _error, ref _warning);
        _swModel = (ModelDoc2)_swApp.ActiveDoc;
        _swDrawing = (DrawingDoc)_swModel;
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
}
