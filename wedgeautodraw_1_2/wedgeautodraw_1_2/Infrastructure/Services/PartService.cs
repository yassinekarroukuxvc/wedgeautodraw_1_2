using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services;

public class PartService : IPartService
{
    private readonly SldWorks _swApp;
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
            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "",
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

            if (!eqMgr.UpdateValuesFromExternalEquationFile())
                Logger.Warn("Failed to update values from the external equation file.");

            eqMgr.AutomaticRebuild = true;
            eqMgr.AutomaticSolveOrder = true;
        }
        catch (Exception ex)
        {
            Logger.Error("Error updating equations: " + ex.Message);
        }
    }

    public void SetEngravedText(string text)
    {
        try
        {
            _custPropMgr = _swModel.Extension.get_CustomPropertyManager("");
            _custPropMgr.Set2("Engraved Text", text);
        }
        catch (Exception ex)
        {
            Logger.Error("Error setting engraved text: " + ex.Message);
        }
    }

    public void ApplyTolerances(NamedDimensionValues dimensions)
    {
        try
        {
            var targets = new[]
            {
                new { Name = "TL", Sketch = "TL_cutting" },
                new { Name = "TD", Sketch = "sketch_TL_cutting" },
                new { Name = "TDF", Sketch = "sketch_TDF_grinding" },
                new { Name = "W", Sketch = "sketch_ISA_grinding" },
                new { Name = "FL", Sketch = "sketch_FA_BA_grinding" }
            };

            foreach (var target in targets)
            {
                _swModel.ClearSelection2(true);
                if (!_swModelExt.SelectByID2($"{target.Name}@{target.Sketch}", "DIMENSION", 0, 0, 0, false, 0, null, 0))
                {
                    Logger.Warn($"Failed to select dimension {target.Name}@{target.Sketch}");
                    continue;
                }

                var selectionMgr = (ISelectionMgr)_swModel.SelectionManager;
                if (selectionMgr.GetSelectedObject6(1, 0) is not DisplayDimension dispDim)
                {
                    Logger.Warn($"Failed to cast selected object to DisplayDimension for {target.Name}");
                    continue;
                }

                dispDim.MarkedForDrawing = true;
                var tol = dispDim.GetDimension2(0).Tolerance;

                double upper = dimensions[target.Name].GetTolerance(Unit.Meter, "+");
                double lower = dimensions[target.Name].GetTolerance(Unit.Meter, "-");

                tol.Type = (upper != lower) ? (int)swTolType_e.swTolBILAT : (int)swTolType_e.swTolSYMMETRIC;
                tol.SetValues(-lower, upper);
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error applying tolerances: {ex.Message}");
        }
    }

    public void ToggleSketchVisibility(string sketchName, bool visible)
    {
        try
        {
            _swModel.ClearSelection2(true);
            if (!_swModelExt.SelectByID2(sketchName, "SKETCH", 0, 0, 0, false, 0, null, 0))
            {
                Logger.Warn($"Failed to select sketch '{sketchName}'.");
                return;
            }

            if (visible)
                _swModel.UnblankSketch();
            else
                _swModel.BlankSketch();
        }
        catch (Exception ex)
        {
            Logger.Error($"Error toggling sketch visibility: {ex.Message}");
        }
    }

    public void EnableSolveOrder(bool enable) => _swModel.GetEquationMgr().AutomaticSolveOrder = enable;

    public void EnableAutoRebuild(bool enable) => _swModel.GetEquationMgr().AutomaticRebuild = enable;

    public void Rebuild()
    {
        try
        {
            _swModel.ForceRebuild3(false);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error during rebuild: {ex.Message}");
        }
    }

    public void Save(bool close = false)
    {
        try
        {
            _swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref _error, ref _warning);
            if (close)
                _swApp.CloseDoc(_partPath);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error during save: {ex.Message}");
        }
    }

    public void Reopen(string partPath)
    {
        _partPath = partPath;
        _swApp.OpenDoc6(
            partPath,
            (int)swDocumentTypes_e.swDocPART,
            (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel,
            "",
            ref _error,
            ref _warning);

        _swModel = (ModelDoc2)_swApp.ActiveDoc;

        if (_swModel == null)
        {
            Logger.Warn("Failed to reopen the part. Model is null.");
            return;
        }

        _swModelExt = _swModel.Extension;
        _swModel.Lock();
    }

    public void Unlock()
    {
        try
        {
            _swModel.UnLock();
        }
        catch (Exception ex)
        {
            Logger.Error("Error unlocking part: " + ex.Message);
        }
    }
    public ModelDoc2 GetModel()
    {
        return _swModel;
    }

}
