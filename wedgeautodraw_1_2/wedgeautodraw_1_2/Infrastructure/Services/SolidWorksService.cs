using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Interfaces;

namespace wedgeautodraw_1_2.Infrastructure.Services;

public class SolidWorksService : ISolidWorksService
{
    private SldWorks _swApp;

    public SldWorks GetApplication(bool visible = true, bool userControl = true, bool backgroundControl = true)
    {
        const string progId = "SldWorks.Application";

        try
        {
            _swApp = new SldWorks();
        }
        catch
        {
            _swApp = Activator.CreateInstance(Type.GetTypeFromProgID(progId)) as SldWorks;
        }

        _swApp.Visible = visible;
        _swApp.UserControl = userControl;
        _swApp.UserControlBackground = backgroundControl;


        return _swApp;
    }
    public void CloseApplication()
    {
        try
        {
            if (_swApp != null)
            {
                _swApp.ExitApp();
                _swApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error closing SolidWorks: {ex.Message}");
        }
    }
}
