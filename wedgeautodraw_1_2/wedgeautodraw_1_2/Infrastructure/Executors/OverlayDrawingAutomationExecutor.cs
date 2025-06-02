using ClosedXML.Excel;
using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;
using wedgeautodraw_1_2.Infrastructure.Services;

namespace wedgeautodraw_1_2.Infrastructure.Executors
{
    class OverlayDrawingAutomationExecutor : IDrawingAutomationExecutor
    {
        public void Run(
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
            string outputTiffPath)
        {
            Logger.Info("=== Starting Overlay Drawing Automation ===");

            var drawingService = InitializeDrawing(swApp, partPath, drawingPath, modPartPath, modDrawingPath);
            UpdateViewScalesAndPositions(drawingService, drawingData,wedgeData);
            FinalizeDrawing(drawingService, partService, outputPdfPath,outputTiffPath);

            Logger.Success("Overlay drawing automation completed.");
        }

        private static DrawingService InitializeDrawing(
            SldWorks swApp,
            string partPath,
            string drawingPath,
            string modPartPath,
            string modDrawingPath)
        {
            var drawingService = new DrawingService(swApp);
            drawingService.ReplaceReferencedModel(modDrawingPath, partPath, modPartPath);
            drawingService.OpenDrawing(modDrawingPath);
            drawingService.Rebuild();

            Logger.Info("Drawing initialized and rebuilt.");
            return drawingService;
        }

        private static void UpdateViewScalesAndPositions(IDrawingService drawingService, DrawingData drawData,WedgeData wedgeData)
        {
            var model = drawingService.GetModel();
            string[] viewNames = new[]
            {
                Constants.FrontView,
                Constants.SideView,
                Constants.TopView,
                Constants.OverlayDetailView,
                Constants.OverlaySectionView
            };

            foreach (var viewName in viewNames)
            {
                var view = new ViewService(viewName, ref model);
                if (viewName == Constants.OverlayFrontView || viewName == Constants.OverlaySideView)
                {
                    view.SetViewScale(3);
                }

                if (viewName == Constants.OverlayDetailView || viewName == Constants.OverlaySectionView)
                {
                    view.SetViewScale(100);
                    view.SetOverlayBreaklineRightShift(0.030);
                   /* view.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
                    {
                        { "ISA", "SelectByName" },
                    });*/   
                }
                if (viewName == Constants.OverlayDetailView)
                {
                    // Read tolerances from WedgeData
                    double W_lowerTol = wedgeData.Dimensions["W"].GetTolerance(Unit.Meter, "-");
                    double W_upperTol = wedgeData.Dimensions["W"].GetTolerance(Unit.Meter, "+");

                    double FL_lowerTol = wedgeData.Dimensions["FL"].GetTolerance(Unit.Meter, "-");
                    double FL_upperTol = wedgeData.Dimensions["FL"].GetTolerance(Unit.Meter, "+");

                    Logger.Success($"W -> Upper : {W_upperTol} - Lower : {W_lowerTol} ");
                    Logger.Success($"FL -> Upper : {FL_upperTol} - Lower : {FL_lowerTol} ");

                    view.SetSketchDimensionValue("D1@Sketch100", W_upperTol);
                    view.SetSketchDimensionValue("D2@Sketch100", W_lowerTol);

                    view.SetSketchDimensionValue("D1@Sketch231", FL_upperTol);  
                    view.SetSketchDimensionValue("D2@Sketch231", FL_lowerTol);
                    view.MoveViewToPosition(-1.615013, -0.015391);
                    //view.DrawBorderBoxInView();
                    //view.DrawTestBox();
                    //view.SetOverlayBreaklinePosition(wedgeData.Dimensions, drawData);
                    //view.InsertModelDimensioning();
                    //view.SetOverlayBreaklineRightShift(0.030);
                }



            }
        }

        private static void FinalizeDrawing(IDrawingService drawingService, IPartService partService, string outputPdfPath,string outputTiffPath)
        {
            var model = drawingService.GetModel();

            model.GraphicsRedraw2();
           //drawingService.SaveAsPdfAndConvertToTiff(outputPdfPath, outputTiffPath);
            drawingService.SaveAsTiff(outputPdfPath);
            drawingService.Unlock();
            drawingService.SaveAndCloseDrawing();

            partService.Save(close: true);
        }
    }
}
