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
        string tolerancePath = Path.Combine(resourcePath, "tolerances.txt");
        string configurationPath = Path.Combine(resourcePath, "configurations.txt");
        string excelPath = Path.Combine(resourcePath, "Copy.xlsx");

        ISolidWorksService swService = new SolidWorksService();
        SldWorks swApp = swService.GetApplication();

        var excelLoader = new ExcelWedgeDataLoader(excelPath);
        var wedgeDataList = excelLoader.LoadWedgeDataList();
        var drawingDataList = excelLoader.LoadDrawingDataList();

        for (int i = 0; i < wedgeDataList.Count; i++)
        {
            var wedge = wedgeDataList[i];
            var drawing = drawingDataList[i];


            string wedgeId = wedge.Metadata.ContainsKey("drawing_number") ? wedge.Metadata["drawing_number"].ToString() : $"Wedge_{i + 1}";
            string outputDir = Path.Combine(resourcePath, "Generated", wedgeId);
            Directory.CreateDirectory(outputDir);

            string modPartPath = Path.Combine(outputDir, $"{wedgeId}.SLDPRT");
            string modDrawingPath = Path.Combine(outputDir, $"{wedgeId}.SLDDRW");
            string modEquationPath = Path.Combine(outputDir, $"equations.txt");
            string outputPdfPath = Path.Combine(outputDir, $"{wedgeId}.pdf");

            FileHelper.CopyTemplateFile(partPath, modPartPath);
            FileHelper.CopyTemplateFile(drawingPath, modDrawingPath);
            FileHelper.CopyTemplateFile(equationPath, modEquationPath);

            IDataContainerLoader dataLoader = new DataContainerLoader(modEquationPath);
            var fullDrawingData = dataLoader.LoadDrawingData(wedge, configurationPath);

            EquationFileUpdater.UpdateEquationFile(modEquationPath, wedge);
            //AutomationExecutor.RunPartAutomation(swApp, partPath, modEquationPath, modPartPath, modEquationPath, wedge);
            //AutomationExecutor.RunDrawingAutomation(swApp, drawingPath, modDrawingPath, partPath, modPartPath, fullDrawingData, wedge, outputPdfPath);
            AutomationExecutor.RunPartAndDrawingAutomation(
                swApp,
                partPath,
                drawingPath,
                modPartPath,
                modDrawingPath,
                modEquationPath,
                fullDrawingData,
                wedge,
                outputPdfPath
            );
            Console.WriteLine($"✅ Processed wedge: {wedgeId}");
        }

        Console.WriteLine("=== DONE ===");
    }
}