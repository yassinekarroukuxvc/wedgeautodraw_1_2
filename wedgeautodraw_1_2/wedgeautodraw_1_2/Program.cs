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
        // === CONFIGURATION ===
        private const bool UseSections = false; // If true, run only one section; if false, run all
        private const int SectionToRun = 3;     // 1, 2, or 3 when UseSections is true

        static void Main(string[] args)
        {
            Console.WriteLine("=== SolidWorks Drawing Automation ===");

            Stopwatch stopwatch = Stopwatch.StartNew();
            ProcessHelper.KillAllSolidWorksProcesses();

            string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
            string resourcePath = Path.Combine(projectRoot, "Resources", "Templates");

            string prodPartTemplate = Path.Combine(resourcePath, "wedge.SLDPRT");
            prodPartTemplate = "C:\\Users\\mounir\\Desktop\\TestTest\\Final\\wedge1.SLDPRT";
            string prodDrawingTemplate = Path.Combine(resourcePath, "wedge.SLDDRW");
            prodDrawingTemplate = "C:\\Users\\mounir\\Desktop\\TestTest\\Final\\wedge1.SLDDRW";
            string overlayPartTemplate = "C:\\Users\\mounir\\Desktop\\Oussama_test\\overlay_wedgev2.SLDPRT";
            string overlayDrawingTemplate = "C:\\Users\\mounir\\Desktop\\overlay_wedgev4.SLDDRW";

            string equationPath = Path.Combine(resourcePath, "equations.txt");
            string configurationPath = Path.Combine(resourcePath, "drawing_config.json");
            string excelPath = Path.Combine(resourcePath, "CKVD_Data_SG.xlsx");
            string rulePath = Path.Combine(resourcePath, "drawing_rules.json");

            ISolidWorksService swService = new SolidWorksService();
            SldWorks swApp = swService.GetApplication();

            bool runAllWedges = false;
            /*List<string> selectedWedgeIds = new List<string> {
                "22000175", "22000113", "22101937", "22102117", "2017407", "22101937", "22102018", "2022281", "2022282",
                "2022463","2023126","2024389","2024872","2024887","2024895","2024896","2025708","2027736","2026591",
                "2026663","2026657","2027723","2027638","2027483","2027944","2028350","2026402","2028071","2026663",
                "2028501","2028502","2028500","2028503","2028608","2028609","2028809","2028720","2028245","2028923",
                "2029055","2029057","2029063","2029269","2028243","2028242","2029737","2029606","2029739","2029738",
                "2028715","2028717","2030607","2024887","2026989","2031630"
            };*/
            List<string> selectedWedgeIds = new List<string> {
                "22000175","22000113","2032088",
                "2031500","2032444","2028343","2028245","2032277","2033252",
                "2033624","2028342","2028343","2028245","2031499","2036477",
                "2036482","2025257"
            };
            var excelLoader = new ExcelWedgeDataLoader(excelPath);
            var allEntries = excelLoader.LoadAllEntries();

            var filteredEntries = runAllWedges
                ? allEntries
                : allEntries.Where(e => selectedWedgeIds.Contains(e.Wedge.Metadata["drawing_number"].ToString())).ToList();

            Logger.Warn($"Total Number Of Excel Rows : {filteredEntries.Count}");

            List<(WedgeData Wedge, DrawingData Drawing)> entriesToProcess;

            if (UseSections)
            {
                int total = filteredEntries.Count;
                int sectionSize = (int)Math.Ceiling(total / 3.0);

                var section1 = filteredEntries.Take(sectionSize).ToList();
                var section2 = filteredEntries.Skip(sectionSize).Take(sectionSize).ToList();
                var section3 = filteredEntries.Skip(sectionSize * 2).ToList();

                entriesToProcess = SectionToRun switch
                {
                    1 => section1,
                    2 => section2,
                    3 => section3,
                    _ => throw new ArgumentOutOfRangeException(nameof(SectionToRun), "SectionToRun must be 1, 2, or 3")
                };

                Logger.Warn($"Running Section {SectionToRun} with {entriesToProcess.Count} entries.");
            }
            else
            {
                entriesToProcess = filteredEntries;
                Logger.Warn($"Running ALL entries ({entriesToProcess.Count})");
            }

            int wedgeCounter = 0;

            foreach (var entry in entriesToProcess)
            {
                var wedge = entry.Wedge;
                var drawing = entry.Drawing;

                string wedgeId = wedge.Metadata.ContainsKey("drawing_number")
                    ? wedge.Metadata["drawing_number"].ToString()
                    : $"Wedge_{Guid.NewGuid()}";

                DrawingType selectedType = DrawingType.Production;

                string baseOutputDir = @"D:\Generated";
                string outputDir = Path.Combine(baseOutputDir, wedgeId);
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

                GC.Collect();
                GC.WaitForPendingFinalizers();
                swApp.CloseAllDocuments(true);
                ClearSolidWorksTempFiles();

                wedgeCounter++;
            }

            stopwatch.Stop();
            swService.CloseApplication();
            Console.WriteLine($"=== DONE: Total Execution Time = {stopwatch.Elapsed.TotalSeconds:F2} seconds ===");
        }

        private static void ClearSolidWorksTempFiles()
        {
            try
            {
                string tempPath = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "Temp");
                var swTempDirs = Directory.GetDirectories(tempPath, "swx*");
                foreach (var dir in swTempDirs)
                {
                    try
                    {
                        Directory.Delete(dir, true);
                        Logger.Warn("SW* FILES DELETED");
                    }
                    catch
                    {
                        Logger.Warn("SW* FILES NOT DELETED");
                    }
                }

                var dhTempDirs = Directory.GetDirectories(tempPath, "dh*");
                foreach (var dir in dhTempDirs)
                {
                    try
                    {
                        Directory.Delete(dir, true);
                        Logger.Warn("DH FILES DELETED");
                    }
                    catch
                    {
                        Logger.Warn("DH FILES NOT DELETED");
                    }
                }

                string swLocalPath = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "SolidWorks");
                if (Directory.Exists(swLocalPath))
                {
                    var backupDirs = Directory.GetDirectories(swLocalPath, "Backup", SearchOption.AllDirectories);
                    foreach (var dir in backupDirs)
                    {
                        try
                        {
                            Directory.Delete(dir, true);
                            Logger.Warn("TEMP FILES DELETED");
                        }
                        catch
                        {
                            Logger.Warn("TEMP FILES NOT DELETED");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Warning] Could not fully clear temp files: {ex.Message}");
            }
        }
    }
}
