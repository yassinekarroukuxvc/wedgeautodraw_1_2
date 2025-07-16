using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Infrastructure.Services;
using wedgeautodraw_1_2.Infrastructure.Utilities;
using wedgeautodraw_1_2.Infrastructure.Helpers;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Infrastructure.Executors;
using wedgeautodraw_1_2.Core.Models;
using SolidWorks.Interop.swconst;

namespace wedgeautodraw_1_2
{
    class Program
    {
        private const bool UseSections = false;
        private const int SectionToRun = 1;

        static void Main(string[] args)
        {
            Console.WriteLine("=== SolidWorks Drawing Automation ===");

            Stopwatch stopwatch = Stopwatch.StartNew();
            ProcessHelper.KillAllSolidWorksProcesses();

            // === Select wedge type ===
            WedgeType selectedWedgeType = WedgeType.COB;
            var paths = PrepareTemplatePaths(selectedWedgeType);

            var excelLoader = new ExcelWedgeDataLoader(paths.ExcelPath, selectedWedgeType);
            var allEntries = excelLoader.LoadAllEntries();

            var filteredEntries = FilterEntries(allEntries, runAll: true);
            var entriesToProcess = SelectSectionsIfNeeded(filteredEntries);

            Logger.Warn($"Total entries to process: {entriesToProcess.Count}");

            ISolidWorksService swService = new SolidWorksService();
            SldWorks swApp = swService.GetApplication();

            ProcessEntries(entriesToProcess, swApp, paths, selectedWedgeType);

            stopwatch.Stop();
            swService.CloseApplication();

            Console.WriteLine($"=== DONE: Total Execution Time = {stopwatch.Elapsed.TotalSeconds:F2} seconds ===");
        }

        private static (string ProdPartTemplate, string ProdDrawingTemplate, string OverlayPartTemplate, string OverlayDrawingTemplate, string ExcelPath, string EquationPath, string ConfigPath, string RulePath) PrepareTemplatePaths(WedgeType wedgeType)
        {
            string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
            string resourcePath = Path.Combine(projectRoot, "Resources", "Templates");

            switch (wedgeType)
            {
                case WedgeType.CKVD:
                    return (
                        ProdPartTemplate: @"C:\Users\mounir\Desktop\TestTest\Final\wedge1.SLDPRT",
                        ProdDrawingTemplate: @"C:\Users\mounir\Desktop\TestTest\Final\wedge1.SLDDRW",
                        OverlayPartTemplate: @"C:\Users\mounir\Desktop\Oussama_test\overlay_wedgev2.SLDPRT",
                        OverlayDrawingTemplate: @"C:\Users\mounir\Desktop\overlay_wedgev4.SLDDRW",
                        ExcelPath: @"C:\Users\mounir\Desktop\wedgeautodraw_1_2\wedgeautodraw_1_2\wedgeautodraw_1_2\Resources\Templates\CKVD_DATA_10_Parts.xlsx",
                        EquationPath: Path.Combine(resourcePath, "equations.txt"),
                        ConfigPath: Path.Combine(resourcePath, "drawing_config.json"),
                        RulePath: Path.Combine(resourcePath, "drawing_rules.json")
                    );

                case WedgeType.COB:
                    return (
                        ProdPartTemplate: Path.Combine(resourcePath, "COB", "mod_wedge.SLDPRT"),
                        ProdDrawingTemplate: Path.Combine(resourcePath, "COB", "mod_wedge.SLDDRW"),
                        OverlayPartTemplate: Path.Combine(resourcePath, "COB", "overlay_wedgev2.SLDPRT"),
                        OverlayDrawingTemplate: Path.Combine(resourcePath, "COB", "overlay_wedgev4.SLDDRW"),
                        ExcelPath: Path.Combine(resourcePath, "COB", "UT-US-COB_Produced_19-23_SG.xlsx"),
                        EquationPath: Path.Combine(resourcePath, "COB", "equations1.txt"),
                        ConfigPath: Path.Combine(resourcePath, "drawing_config.json"),
                        RulePath: Path.Combine(resourcePath, "drawing_rules.json")
                    );

                default:
                    throw new NotSupportedException($"Paths for wedge type '{wedgeType}' not configured.");
            }
        }

        private static List<(WedgeData Wedge, DrawingData Drawing)> FilterEntries(List<(WedgeData Wedge, DrawingData Drawing)> allEntries, bool runAll)
        {
            List<string> selectedIds = new()
            {
                "14007448","14007090","14004692","14004692","14009907"
            };

            return runAll
                ? allEntries
                : allEntries.Where(e => selectedIds.Contains(e.Wedge.Metadata["drawing_number"].ToString())).ToList();
        }

        private static List<(WedgeData Wedge, DrawingData Drawing)> SelectSectionsIfNeeded(List<(WedgeData Wedge, DrawingData Drawing)> entries)
        {
            if (!UseSections)
            {
                Logger.Warn($"Running ALL entries ({entries.Count})");
                return entries;
            }

            int total = entries.Count;
            int sectionSize = (int)Math.Ceiling(total / 3.0);

            var section1 = entries.Take(sectionSize).ToList();
            var section2 = entries.Skip(sectionSize).Take(sectionSize).ToList();
            var section3 = entries.Skip(sectionSize * 2).ToList();

            var selectedSection = SectionToRun switch
            {
                1 => section1,
                2 => section2,
                3 => section3,
                _ => throw new ArgumentOutOfRangeException(nameof(SectionToRun), "SectionToRun must be 1, 2, or 3")
            };

            Logger.Warn($"Running Section {SectionToRun} with {selectedSection.Count} entries.");
            return selectedSection;
        }

        private static void ProcessEntries(List<(WedgeData Wedge, DrawingData Drawing)> entries, SldWorks swApp, (string ProdPartTemplate, string ProdDrawingTemplate, string OverlayPartTemplate, string OverlayDrawingTemplate, string ExcelPath, string EquationPath, string ConfigPath, string RulePath) paths, WedgeType wedgeType)
        {
            int counter = 0;

            foreach (var entry in entries)
            {
                var wedge = entry.Wedge;
                var drawing = entry.Drawing;

                string wedgeId = wedge.Metadata.ContainsKey("drawing_number")
                    ? wedge.Metadata["drawing_number"].ToString()
                    : $"Wedge_{Guid.NewGuid()}";

                DrawingType selectedType = DrawingType.Production;

                string baseOutputDir = $@"D:\{wedgeType}";
                string wedgeDir = Path.Combine(baseOutputDir, wedgeId);
                string typeFolder = selectedType == DrawingType.Production ? "Production" : "Overlay";
                string outputDir = Path.Combine(wedgeDir, typeFolder);
                Directory.CreateDirectory(outputDir);

                string templatePartPath = selectedType == DrawingType.Production ? paths.ProdPartTemplate : paths.OverlayPartTemplate;
                string templateDrawingPath = selectedType == DrawingType.Production ? paths.ProdDrawingTemplate : paths.OverlayDrawingTemplate;

                string modEquationPath = Path.Combine(outputDir, "equations.txt");
                string modPartPath = Path.Combine(outputDir, $"{(selectedType == DrawingType.Overlay ? "overlay_" : "")}{wedgeId}.SLDPRT");
                string modDrawingPath = Path.Combine(outputDir, $"{(selectedType == DrawingType.Overlay ? "overlay_" : "")}{wedgeId}.SLDDRW");
                string outputPdfPath = Path.Combine(outputDir, $"{(selectedType == DrawingType.Overlay ? "overlay_" : "")}{wedgeId}.pdf");
                string outputTiffPath = selectedType == DrawingType.Overlay ? Path.Combine(outputDir, $"{wedgeId}.tif") : "";

                FileHelper.CopyTemplateFile(templatePartPath, modPartPath);
                FileHelper.CopyTemplateFile(templateDrawingPath, modDrawingPath);
                FileHelper.CopyTemplateFile(paths.EquationPath, modEquationPath);

                EquationFileUpdater.UpdateEquationFile(modEquationPath, wedge);

                IDrawingDataLoader dataLoader = new DrawingDataLoader(modEquationPath);
                var fullDrawingData = dataLoader.LoadDrawingData(wedge, paths.ConfigPath, selectedType);

                var ruleEngine = new DrawingRuleEngine(paths.RulePath);
                ruleEngine.Apply(fullDrawingData, wedge, selectedType);

                DynamicDimensionStyler.ApplyDynamicStyles(fullDrawingData, wedge);

                var partService = PartAutomationExecutor.Run(swApp, modEquationPath, modPartPath, wedge);

                DrawingAutomationExecutor.Run(
                    swApp,
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

                GC.Collect();
                GC.WaitForPendingFinalizers();
                swApp.CloseAllDocuments(true);
                ClearSolidWorksTempFiles();

                counter++;
            }
        }

        private static void ClearSolidWorksTempFiles()
        {
            try
            {
                string tempPath = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "Temp");
                var swTempDirs = Directory.GetDirectories(tempPath, "swx*");
                var dhTempDirs = Directory.GetDirectories(tempPath, "dh*");

                DeleteDirs(swTempDirs, "SW* FILES");
                DeleteDirs(dhTempDirs, "DH FILES");

                string swLocalPath = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "SolidWorks");
                if (Directory.Exists(swLocalPath))
                {
                    var backupDirs = Directory.GetDirectories(swLocalPath, "Backup", SearchOption.AllDirectories);
                    DeleteDirs(backupDirs, "TEMP FILES");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Warning] Could not fully clear temp files: {ex.Message}");
            }
        }

        private static void DeleteDirs(string[] dirs, string label)
        {
            foreach (var dir in dirs)
            {
                try
                {
                    Directory.Delete(dir, true);
                    Logger.Warn($"{label} DELETED");
                }
                catch
                {
                    Logger.Warn($"{label} NOT DELETED");
                }
            }
        }
    }
}
