using ClosedXML.Excel;
using SolidWorks.Interop.sldworks;
using System.Transactions;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Factories;
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
            UpdateViewScalesAndPositions(swApp, drawingService, drawingData, wedgeData);
            var noteService = new NoteService(swApp,drawingService.GetModel());

            // Insert dimension note as an alternative or supplement
            noteService.InsertDimensionNote(
                drawingData.TablePositions[Constants.DimensionTable],
                drawingData.DimensionKeysInTable,
                "DIMENSIONS:",
                drawingData,
                wedgeData.Dimensions
            );
            

            FinalizeDrawing(drawingService, partService, outputPdfPath, outputTiffPath,wedgeData,noteService);

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
            drawingService.Unlock();
            swApp.Visible = true;
            var conf = new TiffExportSettings(swApp);
            //conf.RunSolidWorksMacro(swApp, "C:\\Users\\mounir\\Desktop\\wedgeautodraw_1_2\\wedgeautodraw_1_2\\wedgeautodraw_1_2\\Resources\\Templates\\Macro.swp");
            conf.SetTiffExportSettings();
            //System.Threading.Thread.Sleep(3000);
            drawingService.Lock();
            drawingService.Rebuild();

            Logger.Info("Drawing initialized and rebuilt.");
            return drawingService;
        }

        private static void UpdateViewScalesAndPositions(SldWorks swApp, IDrawingService drawingService, DrawingData drawData, WedgeData wedgeData)
        {
            var model = drawingService.GetModel();
            double W_lowerTol = wedgeData.Dimensions["W"].GetTolerance(Unit.Meter, "-");
            double W_upperTol = wedgeData.Dimensions["W"].GetTolerance(Unit.Meter, "+");

            double FL_lowerTol = wedgeData.Dimensions["FL"].GetTolerance(Unit.Meter, "-");
            double FL_upperTol = wedgeData.Dimensions["FL"].GetTolerance(Unit.Meter, "+");
            double b= wedgeData.Dimensions["B"].GetTolerance(Unit.Meter, "+");
            Logger.Info($"Tolerances for W: -{W_lowerTol:F4} / +{W_upperTol:F4}");
            Logger.Info($"Tolerances for FL: -{FL_lowerTol:F4} / +{FL_upperTol:F4}");
            Logger.Info($"Tolerance for B (Upper only): +{b:F4}");
            string[] viewNames = new[]
            {
                Constants.OverlaySideView,
                Constants.OverlayTopView,
                Constants.OverlayDetailView,
                Constants.OverlaySectionView
            };

            foreach (var viewName in viewNames)
            {
                Logger.Info($"Processing view: {viewName}");

                //var view = new ViewService(viewName, ref model);
                var viewFactory = new StandardViewFactory(model);
                var view = viewFactory.CreateView(viewName);

                if (viewName == Constants.OverlayDetailView || viewName == Constants.OverlaySectionView)
                {
                    double fl = wedgeData.Dimensions["FL"].GetValue(Unit.Millimeter);
                    double scale = wedgeData.OverlayScaling;
                    view.SetViewScale(scale);
                }

                if (viewName == Constants.OverlaySideView)
                {
                    if (wedgeData.Dimensions["TL"].GetValue(Unit.Millimeter) < 40)
                    {
                        view.SetViewScale(4);
                    }

                    view.CreateCenterline(wedgeData.Dimensions, drawData);
                }

                if (viewName == Constants.OverlayTopView)
                {
                    if (wedgeData.Dimensions["TL"].GetValue(Unit.Millimeter) < 40)
                    {
                        view.SetViewScale(4);
                    }
                    view.CreateCentermark(wedgeData.Dimensions, drawData);
                }

                if (viewName == Constants.OverlaySectionView)
                {
                    view.SetOverlayBreaklinePosition(wedgeData.Dimensions, drawData);
                    view.CenterViewVertically();
                    view.SetBreakLineGap(0.000025);
                    view.AlignViewHorizontally(false);
                    view.CenterSectionViewVisuallyVertically(wedgeData.Dimensions);
                    view.SetSketchDimensionValue("D1@Sketch474", W_upperTol);
                    view.SetSketchDimensionValue("D3@Sketch474", W_lowerTol);
                }

                if (viewName == Constants.OverlayDetailView)
                {
                    view.SetOverlayBreaklinePosition(wedgeData.Dimensions, drawData);
                    view.CenterViewVertically();
                    view.SetBreakLineGap(0.000025);
                    view.AlignViewHorizontally(true);
                    view.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
                    {
                        { "ISA", "SelectByName" },
                        { "GA", "SelectByName" },
                    },drawData.DrawingType);
                    view.SetSketchDimensionValue("D1@Sketch447", FL_upperTol);
                    view.SetSketchDimensionValue("D2@Sketch447", FL_lowerTol);
                    view.SetSketchDimensionValue("D3@Sketch447", b);
                }

                view = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void FinalizeDrawing(IDrawingService drawingService, IPartService partService, string outputPdfPath, string outputTiffPath,WedgeData wedgedata,INoteService noteService)
        {
            var model = drawingService.GetModel();

            model.GraphicsRedraw2();
            drawingService.ZoomToSheet();
            double overlay_calibration = (double.Parse(wedgedata.OverlayCalibration) / 25400) * wedgedata.OverlayScaling ;
            drawingService.DrawCenteredSquareOnSheet(overlay_calibration);
            noteService.InsertOverlayCalibrationNote(wedgedata.OverlayCalibration,overlay_calibration);
            drawingService.SaveAsTiff(outputTiffPath);
            drawingService.SaveAsPdf(outputPdfPath);

            drawingService.Unlock();
            drawingService.SaveAndCloseDrawing();

            partService.Save(close: true);
        }
    }
}
