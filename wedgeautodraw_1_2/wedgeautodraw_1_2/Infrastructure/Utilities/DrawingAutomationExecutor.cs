using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Factories;
using wedgeautodraw_1_2.Infrastructure.Helpers;
using wedgeautodraw_1_2.Infrastructure.Services;
using wedgeautodraw_1_2.Infrastructure.Executors;

namespace wedgeautodraw_1_2.Infrastructure.Utilities;

public static class DrawingAutomationExecutor
{
    public static void Run(
        SldWorks swApp,
        IPartService partService,
        string partPath,
        string drawingPath,
        string modPartPath,
        string modDrawingPath,
        string modEquationPath,
        DrawingData drawingData,
        WedgeData wedgeData,
        string outputPdfPath,
        string outputTiffPath,
        DrawingType type)
    {
        Logger.Info("=== Starting Drawing Automation ===");
        //Console.WriteLine(drawingData);
        Console.WriteLine(wedgeData);
        IDrawingAutomationExecutor executor = type switch
        {
            DrawingType.Production => new ProductionDrawingAutomationExecutor(),
            DrawingType.Overlay => new OverlayDrawingAutomationExecutor(),
            _ => throw new NotSupportedException($"DrawingType {type} not supported")
        };

        executor.Run(swApp, partService, partPath, drawingPath, modPartPath, modDrawingPath, modEquationPath, drawingData, wedgeData, outputPdfPath,outputTiffPath);


        Logger.Success("Drawing automation completed.");
    }
}
