
using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;
using wedgeautodraw_1_2.Infrastructure.Services;
using wedgeautodraw_1_2.Core.Enums;
namespace wedgeautodraw_1_2.Infrastructure.Utilities;

public static class PartAutomationExecutor
{
    public static IPartService Run(SldWorks swApp, string equationPath, string partPath, WedgeData wedge)
    {
        var partService = new PartService(swApp);
        partService.OpenPart(partPath);
        partService.ApplyTolerances(wedge.Dimensions);
        partService.UpdateEquations(equationPath);
        EquationFileUpdater.EnsureAllEquationsExist(partService.GetModel(), wedge);
        //partService.SetEngravedText(wedge.EngravedText);
        // ===== Supress if GD = 0 =======
        if (wedge.Dimensions.ContainsKey("GA"))
        {
            double gaValue = wedge.Dimensions["GA"].GetValue(Unit.Millimeter);
            if (gaValue == 0)
            {
                Logger.Info("GA value is 0. Suppressing feature 'Cut-Extrude1'.");
                partService.SuppressOrUnsuppressSketch("Sketch1", suppress: true);
                partService.SuppressOrUnsuppressFeature("Cut-Extrude1", suppress: true);
            }
            else
            {
                Logger.Info("GA value is not 0. Unsuppressing feature 'Cut-Extrude1'.");
                partService.SuppressOrUnsuppressFeature("Cut-Extrude1", suppress: false);
            }
        }
        else
        {
            Logger.Warn("GA key not found in wedge data. Skipping feature suppression.");
        }

        partService.SetEngravedText("XXXX-XXX-XXX-XX");
        partService.Rebuild();
        partService.Save();
        return partService;
    }
}

