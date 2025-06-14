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
            UpdateViewScalesAndPositions(drawingService, drawingData, wedgeData);
            var tableService = new TableService(swApp, drawingService.GetModel());

            tableService.CreateDimensionNote(drawingData.TablePositions[Constants.DimensionTable], drawingData.DimensionKeysInTable, "DIMENSIONS:", drawingData, wedgeData.Dimensions);

            FinalizeDrawing(drawingService, partService, outputPdfPath, outputTiffPath);
            

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

        private static void UpdateViewScalesAndPositions(IDrawingService drawingService, DrawingData drawData, WedgeData wedgeData)
        {
            var model = drawingService.GetModel();
            string[] viewNames = new[]
            {
                Constants.OverlayFrontView,
                Constants.OverlaySideView,
                Constants.OverlayTopView,
                Constants.OverlayDetailView,
                Constants.OverlaySectionView
            };

            foreach (var viewName in viewNames)
            {
                Logger.Info($"Processing view: {viewName}");

                var view = new ViewService(viewName, ref model);

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
                    // view.SetViewPosition(drawData.ViewPositions[Constants.SideView]);
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
                    /*var centerPosition = view.GetVerticalSheetCenterPosition(drawData);
                    if (centerPosition != null)
                    {
                        double[] values = centerPosition.GetValues(Unit.Millimeter);
                        Logger.Warn($"Center Position: X = {values[0]:F4} mm, Y = {values[1]:F4} mm");

                        view.SetViewPosition(centerPosition);
                        Logger.Info($"Section view vertically centered using calculated position.");
                    }
                    else
                    {
                        Logger.Warn("Failed to calculate vertical center position for section view.");
                    }*/
                    view.CenterViewVertically();
                    view.SetViewX(140);
                    view.CenterSectionViewVisuallyVertically(wedgeData.Dimensions);
                    //view.AlignViewHorizontally(isDetailView: false);
                    //view.AlignViewRightEdgeToTarget(alignToSheetCenter: true);


                }

                if (viewName == Constants.OverlayDetailView)
                {
                    view.SetOverlayBreaklinePosition(wedgeData.Dimensions,drawData);
                    view.CenterViewVertically();
                    //view.AlignViewHorizontally(isDetailView: true);
                    //view.AlignViewRightEdgeToTarget(alignToSheetCenter: false);
                    view.SetViewX(339);
                    
                    //view.ShiftViewRightWithClamping(0.002);
                    /*view.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
                    {
                        { "ISA", "SelectByName" },
                    });*/
                }

                // ====== CRITICAL: Force ViewService cleanup ======
                view = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                // ===============================================
            }
        }

        private static void FinalizeDrawing(IDrawingService drawingService, IPartService partService, string outputPdfPath, string outputTiffPath)
        {
            var model = drawingService.GetModel();

            model.GraphicsRedraw2();
            drawingService.ZoomToSheet();
            drawingService.DrawCenteredSquareOnSheet(4.5);

            // Optional saves
            drawingService.SaveAsTiff(outputTiffPath);
            drawingService.SaveAsPdf(outputPdfPath);

            drawingService.Unlock();
            drawingService.SaveAndCloseDrawing();

            partService.Save(close: true);
        }
    }
}
