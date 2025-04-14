using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Services;
using wedgeautodraw_1_2.Infrastructure.Utilities;
namespace wedgeautodraw_1_2;
class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("=== SolidWorks Drawing Automation ===");

        string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
        string resourcePath = Path.Combine(projectRoot, "Resources", "Templates");

        string partPath = Path.Combine(resourcePath, "wedge.SLDPRT");
        string drawingPath = Path.Combine(resourcePath, "wedge.SLDDRW");
        string equationPath = Path.Combine(resourcePath, "equations.txt");
        string tolerencePath = Path.Combine(resourcePath, "tolerances.txt");
        string configurationPath = Path.Combine(resourcePath, "configurations.txt");

        string modPartPath = Path.Combine(resourcePath, "mod_wedge.SLDPRT");
        string modDrawingPath = Path.Combine(resourcePath, "mod_wedge.SLDDRW");
        string modEquationPath = Path.Combine(resourcePath, "mod_equations.txt");
        string outputPdfPath = Path.Combine(resourcePath, "mod_wedge.pdf");
        FileHelper.CopyTemplateFile(partPath, modPartPath);
        FileHelper.CopyTemplateFile(drawingPath, modDrawingPath);
        FileHelper.CopyTemplateFile(equationPath, modEquationPath);

        ISolidWorksService swService = new SolidWorksService();
        SldWorks swApp = swService.GetApplication();

        IDataContainerLoader dataLoader = new DataContainerLoader(equationPath, tolerencePath);
        WedgeData wedgeData = dataLoader.LoadWedgeData();
        Console.WriteLine(wedgeData);
        DrawingData drawingData = dataLoader.LoadDrawingData(wedgeData, configurationPath);
        Console.WriteLine(drawingData);

        AutomationExecutor.RunPartAutomation(swApp, partPath, equationPath, modPartPath, modEquationPath, wedgeData);
        AutomationExecutor.RunDrawingAutomation(swApp, drawingPath, modDrawingPath, partPath, modPartPath, drawingData, wedgeData, outputPdfPath);

        Console.WriteLine("=== DONE ===");
    }
}
