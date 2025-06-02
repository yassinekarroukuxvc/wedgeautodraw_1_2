using SolidWorks.Interop.sldworks;

namespace wedgeautodraw_1_2.Core.Interfaces;

public interface ISolidWorksService
{
    SldWorks GetApplication(bool visible = true, bool userControl = true, bool backgroundControl = true);
    public void CloseApplication();
}
