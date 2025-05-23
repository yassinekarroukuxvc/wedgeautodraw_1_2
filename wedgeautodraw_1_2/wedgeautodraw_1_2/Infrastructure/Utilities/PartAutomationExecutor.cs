
using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Services;

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
        partService.SetEngravedText("XXXX-XXX-XXX-XX");
        partService.Rebuild();
        partService.Save();
        return partService;
    }
}

