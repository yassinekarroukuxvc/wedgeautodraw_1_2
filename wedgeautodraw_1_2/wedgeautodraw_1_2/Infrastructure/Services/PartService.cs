using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Infrastructure.Services;

public class PartService : IPartService
{
    private SldWorks _swApp;
    private ModelDoc2 _swModel;
    private ModelDocExtension _swModelExt;
    private CustomPropertyManager _custPropMgr;

    private string _partPath;
    private int _error = 0;
    private int _warning = 0;

    public PartService(SldWorks swApp)
    {
        _swApp = swApp;
    }

    public void OpenPart(string partPath)
    {
        _partPath = partPath;
        _swApp.OpenDoc6(partPath, (int)swDocumentTypes_e.swDocPART,
            (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "",
            ref _error, ref _warning);

        _swModel = (ModelDoc2)_swApp.ActiveDoc;
        _swModelExt = _swModel.Extension;
        _swModel.Lock();
    }

    public void UpdateEquations(string equationFilePath)
    {
        try
        {
            var eqMgr = _swModel.GetEquationMgr();

            eqMgr.FilePath = equationFilePath;


            bool updated = eqMgr.UpdateValuesFromExternalEquationFile();
            if (!updated)
                Console.WriteLine("⚠️ Failed to update values from the external equation file.");
            eqMgr.AutomaticRebuild = true;
            eqMgr.AutomaticSolveOrder = true;
            Thread.Sleep(2000);
            _swModel.ForceRebuild3(false);
        }
        catch (Exception ex)
        {
            Console.WriteLine("❌ Error updating equations: " + ex.Message);
        }
    }


    public void SetEngravedText(string text)
    {
        _swModelExt = _swModel.Extension;
        _custPropMgr = _swModelExt.get_CustomPropertyManager("");
        _custPropMgr.Set2("Engraved Text", text);
        _swModel.ForceRebuild3(false);
    }

    public void ApplyTolerances(DynamicDataContainer dimensions)
    {
        try
        {
            _swModelExt = _swModel.Extension;

            var targets = new[]
            {
            new { Name = "TL", Sketch = "TL_cutting" },
            new { Name = "TD", Sketch = "sketch_TL_cutting" },
            new { Name = "TDF", Sketch = "sketch_TDF_grinding" }
        };

            foreach (var target in targets)
            {
                _swModel.ClearSelection2(true);
                bool selected = _swModelExt.SelectByID2($"{target.Name}@{target.Sketch}", "DIMENSION", 0, 0, 0, false, 0, null, 0);
                if (!selected)
                {
                    Console.WriteLine($"⚠️ Failed to select dimension {target.Name}@{target.Sketch}");
                    continue;
                }

                var selectionMgr = (ISelectionMgr)_swModel.SelectionManager;
                var dispDim = selectionMgr.GetSelectedObject6(1, 0) as DisplayDimension;
                if (dispDim == null)
                {
                    Console.WriteLine($"⚠️ Failed to cast selected object to DisplayDimension for {target.Name}");
                    continue;
                }

                dispDim.MarkedForDrawing = true; // ✅ Enable for drawing

                var tol = dispDim.GetDimension2(0).Tolerance;

                double upper = dimensions[target.Name].GetTolerance(Unit.Meter, "+");
                double lower = dimensions[target.Name].GetTolerance(Unit.Meter, "-");

                tol.Type = (upper != lower)
                    ? (int)swTolType_e.swTolBILAT
                    : (int)swTolType_e.swTolSYMMETRIC;

                tol.SetValues(-lower, upper);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error applying tolerances: {ex.Message}");
        }
    }

    public void ToggleSketchVisibility(string sketchName, bool visible)
    {
        try
        {
            if (_swModel == null || _swModelExt == null)
            {
                Console.WriteLine("❌ SolidWorks model or extension is null. Cannot toggle sketch visibility.");
                return;
            }

            bool selected = _swModelExt.SelectByID2(sketchName, "SKETCH", 0, 0, 0, false, 0, null, 0);

            if (!selected)
            {
                Console.WriteLine($"Failed to select sketch '{sketchName}'.");
                return;
            }

            if (visible)
                _swModel.UnblankSketch();
            else
                _swModel.BlankSketch();

            _swModel.ForceRebuild3(false);
            Console.WriteLine($"Sketch '{sketchName}' visibility set to {(visible ? "visible" : "hidden")}.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error toggling sketch visibility: {ex.Message}");
        }
    }


    public void EnableSolveOrder(bool enable)
    {
        _swModel.GetEquationMgr().AutomaticSolveOrder = enable;
    }

    public void EnableAutoRebuild(bool enable)
    {
        _swModel.GetEquationMgr().AutomaticRebuild = enable;
    }

    public void Rebuild()
    {
        _swModel.ForceRebuild3(false);
    }

    public void Save(bool close = false)
    {
        _swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref _error, ref _warning);

        if (close)
            _swApp.CloseDoc(_partPath);
    }

    public void Reopen(string partPath)
    {
        _partPath = partPath;

        var model = _swApp.OpenDoc6(
            partPath,
            (int)swDocumentTypes_e.swDocPART,
            (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel,
            "",
            ref _error,
            ref _warning
        );

        _swModel = (ModelDoc2)model;

        if (_swModel == null)
        {
            Console.WriteLine("Failed to reopen the part. Model is null.");
            return;
        }

        _swModelExt = _swModel.Extension;
        _swModel.Lock();
    }

    public void Unlock()
    {
        _swModel.UnLock();
    }
}
