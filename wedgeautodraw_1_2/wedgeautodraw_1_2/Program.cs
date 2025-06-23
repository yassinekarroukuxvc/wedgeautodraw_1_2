using System;
using System.Diagnostics;
using System.IO;
using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Infrastructure.Services;
using wedgeautodraw_1_2.Infrastructure.Utilities;
using wedgeautodraw_1_2.Infrastructure.Helpers;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Infrastructure.Executors;
using wedgeautodraw_1_2.Core.Models;
using SolidWorks.Interop.swconst;

namespace wedgeautodraw_1_2;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("=== SolidWorks Drawing Automation ===");

        Stopwatch stopwatch = Stopwatch.StartNew();
        ProcessHelper.KillAllSolidWorksProcesses();

        string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
        string resourcePath = Path.Combine(projectRoot, "Resources", "Templates");

        string prodPartTemplate = Path.Combine(resourcePath, "wedge.SLDPRT");
        string prodDrawingTemplate = Path.Combine(resourcePath, "wedge.SLDDRW");
        string overlayPartTemplate = "C:\\Users\\mounir\\Desktop\\Oussama_test\\overlay_wedgev2.SLDPRT";
        //string overlayDrawingTemplate = "C:\\Users\\mounir\\Desktop\\Oussama_test\\overlay_wedgev5.SLDDRW";
        string overlayDrawingTemplate = "C:\\Users\\mounir\\Desktop\\overlay_wedgev4.SLDDRW";

        string equationPath = Path.Combine(resourcePath, "equations.txt");
        string configurationPath = Path.Combine(resourcePath, "drawing_config.json");
        string excelPath = Path.Combine(resourcePath, "CKVD_DATA_10_Parts.xlsx");
        string rulePath = Path.Combine(resourcePath, "drawing_rules.json");

        ISolidWorksService swService = new SolidWorksService();
        SldWorks swApp = swService.GetApplication();

        var excelLoader = new ExcelWedgeDataLoader(excelPath);
        var allEntries = excelLoader.LoadAllEntries();
      
        foreach (var (wedge, drawing) in allEntries)
        {
            string wedgeId = wedge.Metadata.ContainsKey("drawing_number")
                ? wedge.Metadata["drawing_number"].ToString()
                : $"Wedge_{Guid.NewGuid()}";

            DrawingType selectedType = DrawingType.Overlay;

            string outputDir = Path.Combine(resourcePath, "Generated", wedgeId);
            Directory.CreateDirectory(outputDir);

            string templatePartPath = selectedType == DrawingType.Production ? prodPartTemplate : overlayPartTemplate;
            string templateDrawingPath = selectedType == DrawingType.Production ? prodDrawingTemplate : overlayDrawingTemplate;

            string modEquationPath = Path.Combine(outputDir, "equations.txt");

            string modPartPath;
            string modDrawingPath;
            string outputPdfPath;
            string outputTiffPath;
            if (selectedType == DrawingType.Overlay)
            {
                modPartPath = Path.Combine(outputDir, $"overlay_{wedgeId}.SLDPRT");
                modDrawingPath = Path.Combine(outputDir, $"overlay_{wedgeId}.SLDDRW");
                outputPdfPath = Path.Combine(outputDir, $"overlay_{wedgeId}.pdf");
                outputTiffPath = Path.Combine(outputDir, $"overlay_{wedgeId}.tif");
            }
            else
            {
                modPartPath = Path.Combine(outputDir, $"{wedgeId}.SLDPRT");
                modDrawingPath = Path.Combine(outputDir, $"{wedgeId}.SLDDRW");
                outputPdfPath = Path.Combine(outputDir, $"{wedgeId}.pdf");
                outputTiffPath = "";
            }

            FileHelper.CopyTemplateFile(templatePartPath, modPartPath);
            FileHelper.CopyTemplateFile(templateDrawingPath, modDrawingPath);
            FileHelper.CopyTemplateFile(equationPath, modEquationPath);

            EquationFileUpdater.UpdateEquationFile(modEquationPath, wedge);

            IDrawingDataLoader dataLoader = new DrawingDataLoader(modEquationPath);
            var fullDrawingData = dataLoader.LoadDrawingData(wedge, configurationPath, selectedType);
            var ruleEngine = new DrawingRuleEngine(rulePath);
            ruleEngine.Apply(fullDrawingData, wedge);
            DynamicDimensionStyler.ApplyDynamicStyles(fullDrawingData, wedge);
            var partService = PartAutomationExecutor.Run(swApp, modEquationPath, modPartPath, wedge);

            DrawingAutomationExecutor.Run(swApp,
                partService,
                templatePartPath,
                templateDrawingPath,
                modPartPath,
                modDrawingPath,
                modEquationPath,
                fullDrawingData,
                wedge,
                outputPdfPath,
                outputTiffPath,
                selectedType);

            Console.WriteLine($"Processed wedge: {wedgeId}");
        }

        stopwatch.Stop();
        swService.CloseApplication();
        Console.WriteLine($"=== DONE: Total Execution Time = {stopwatch.Elapsed.TotalSeconds:F2} seconds ===");
    }
}
