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

namespace wedgeautodraw_1_2;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("=== SolidWorks Drawing Automation ===");

        // --- START TIMER ---
        Stopwatch stopwatch = Stopwatch.StartNew();

        ProcessHelper.KillAllSolidWorksProcesses();

        string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
        string resourcePath = Path.Combine(projectRoot, "Resources", "Templates");


        string partPath = Path.Combine(resourcePath, "wedge.SLDPRT");
        string drawingPath = Path.Combine(resourcePath, "wedge.SLDDRW");
        string equationPath = Path.Combine(resourcePath, "equations.txt");
        string tolerancePath = Path.Combine(resourcePath, "tolerances.txt");
        string configurationPath = Path.Combine(resourcePath, "drawing_config.json");
        string excelPath = Path.Combine(resourcePath, "CKVD_DATA_10_Parts.xlsx");

        ISolidWorksService swService = new SolidWorksService();
        SldWorks swApp = swService.GetApplication();

        var excelLoader = new ExcelWedgeDataLoader(excelPath);
        var allEntries = excelLoader.LoadAllEntries();

        foreach (var (wedge, drawing) in allEntries)
        {
            string wedgeId = wedge.Metadata.ContainsKey("drawing_number")
                ? wedge.Metadata["drawing_number"].ToString()
                : $"Wedge_{Guid.NewGuid()}";

            string outputDir = Path.Combine(resourcePath, "Generated", wedgeId);
            Directory.CreateDirectory(outputDir);

            string modPartPath = Path.Combine(outputDir, $"{wedgeId}.SLDPRT");
            string modDrawingPath = Path.Combine(outputDir, $"{wedgeId}.SLDDRW");
            string modEquationPath = Path.Combine(outputDir, $"equations.txt");
            string outputPdfPath = Path.Combine(outputDir, $"{wedgeId}.pdf");

            FileHelper.CopyTemplateFile(partPath, modPartPath);
            FileHelper.CopyTemplateFile(drawingPath, modDrawingPath);
            FileHelper.CopyTemplateFile(equationPath, modEquationPath);

            EquationFileUpdater.UpdateEquationFile(modEquationPath, wedge);

            DrawingType selectedType = DrawingType.Production;
            IDrawingDataLoader dataLoader = new DrawingDataLoader(modEquationPath);
            var fullDrawingData = dataLoader.LoadDrawingData(wedge, configurationPath,selectedType);

            var partService = PartAutomationExecutor.Run(swApp, modEquationPath, modPartPath, wedge);

            ///
            DrawingAutomationExecutor.Run(swApp,
                partService,
                partPath,
                drawingPath,
                modPartPath,
                modDrawingPath,
                modEquationPath,
                fullDrawingData,
                wedge,
                outputPdfPath);

            Console.WriteLine($"Processed wedge: {wedgeId}");
        }

        // --- STOP TIMER ---
        stopwatch.Stop();
        Console.WriteLine($"=== DONE: Total Execution Time = {stopwatch.Elapsed.TotalSeconds:F2} seconds ===");
    }
}
