using ClosedXML.Excel;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swdocumentmgr;
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
            noteService.InsertCustomNoteAsTable(wedgeData.Coining, drawingData.TablePositions[Constants.CoiningNote]);



            FinalizeDrawing(swApp,drawingService, partService, outputPdfPath, outputTiffPath,wedgeData,noteService);

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
            //conf.SetTiffExportSettings();
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
            double b = wedgeData.Dimensions["B"].GetTolerance(Unit.Meter, "+");
            double tl = wedgeData.Dimensions["TL"].GetValue(Unit.Meter);
            double TopSideScale = drawData.ViewScales[Constants.SideView].GetValue(Unit.Millimeter);
            Logger.Info($"Tolerances for W: -{W_lowerTol:F4} / +{W_upperTol:F4}");
            Logger.Info($"Tolerances for FL: -{FL_lowerTol:F4} / +{FL_upperTol:F4}");
            Logger.Info($"Tolerance for B (Upper only): +{b:F4}");
            string wedgeId = wedgeData.Metadata.ContainsKey("drawing_number")
                ? wedgeData.Metadata["drawing_number"].ToString()
                : $"Wedge_{Guid.NewGuid()}";

            var specialWedgeIds = new HashSet<string> { "2026582", "2026989", "2030604" };

            if (specialWedgeIds.Contains(wedgeId))
            {
                Logger.Info($"Wedge ID {wedgeId} matches special IDs — hiding Layer 8.");

                var layerMgr = (ILayerMgr)model.GetLayerManager();
                var layerObj = layerMgr.GetLayer("Layer8");

                if (layerObj != null)
                {
                    var layer = (ILayer)layerObj;
                    layer.Visible = false;

                    Logger.Success("Layer 8 successfully hidden.");
                }
                else
                {
                    Logger.Warn("Layer 8 not found; cannot hide.");
                }
            }
            else
            {
                Logger.Info($"Wedge ID {wedgeId} does not require hiding Layer 8.");
            }


            string[] viewNames = new[]
            {
        Constants.OverlaySideView,
        Constants.OverlayTopView,
        Constants.OverlayDetailView,
        Constants.OverlaySectionView,
        Constants.OverlaySideView2
    };

            IViewService sideView = null;
            IViewService topView = null;

            foreach (var viewName in viewNames)
            {
                Logger.Info($"Processing view: {viewName}");

                var viewFactory = new StandardViewFactory(model);
                var view = viewFactory.CreateView(viewName);
                if (viewName == Constants.OverlaySideView2)
                {
                    view.DeleteOverlaySideView2(wedgeData.Dimensions);
                    view.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
                    {
                         { "VR", "SelectByName" },
                         { "VW", "SelectByName" },
                    }, drawData.DrawingType);
                }

                if (viewName == Constants.OverlayDetailView || viewName == Constants.OverlaySectionView)
                {
                    double scale = wedgeData.OverlayScaling;
                    view.SetViewScale(scale);
                }

                if (viewName == Constants.OverlaySideView)
                {
                    sideView = view;
                    view.SetViewScale(drawData.ViewScales[Constants.SideView].GetValue(Unit.Millimeter));
                    view.AlignViewHorizontally(false, tlInMeters: wedgeData.Dimensions["TL"].GetValue(Unit.Meter));
                    view.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
                    {
                         { "FX", "SelectByName" },
                         { "D3", "SelectByName" },
                         { "FA", "SelectByName" },
                         { "BA", "SelectByName" },
                         { "E", "SelectByName" },
                         { "X", "SelectByName" },
                    }, drawData.DrawingType);
                    view.SetSketchDimensionValue("D1@Sketch33", 0.19 / TopSideScale);
                }


                if (viewName == Constants.OverlayTopView)
                {
                    topView = view;
                    view.SetViewScale(drawData.ViewScales[Constants.TopView].GetValue(Unit.Millimeter));
                    //view.CreateCentermark(wedgeData.Dimensions, drawData);
                    view.SetViewPosition(drawData.ViewPositions["Top_view"]);
                    view.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
                    {
                        { "TDF", "SelectByName" },
                    }, drawData.DrawingType);
                    view.SetSketchDimensionValue("D2@Sketch33", 0.006/ TopSideScale);
                    view.SetSketchDimensionValue("D3@Sketch33", 0.015/TopSideScale);
                    double d3 = view.GetSketchDimensionValue("D3@Sketch33");
                    view.SetSketchDimensionValue("D4@Sketch33",d3 /2);
                }

                if (viewName == Constants.OverlaySectionView)
                {
                    view.SetOverlayBreaklinePosition(wedgeData.Dimensions, drawData);
                    view.CenterViewVertically();
                    view.SetBreakLineGap(0.000025);
                    view.AlignViewHorizontally(false, tlInMeters: 0);
                    view.CenterSectionViewVisuallyVertically(wedgeData.Dimensions);
                    view.SetSketchDimensionValue("D1@Sketch474", FL_upperTol);
                    view.SetSketchDimensionValue("D3@Sketch474", FL_lowerTol);
                }

                if (viewName == Constants.OverlayDetailView)
                {
                    view.SetOverlayBreaklinePosition(wedgeData.Dimensions, drawData);
                    view.CenterViewVertically();
                    view.SetBreakLineGap(0.000025);
                    view.AlignViewHorizontally(true, tlInMeters: 0);
                    view.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
                    {
                        { "ISA", "SelectByName" },
                        { "GA", "SelectByName" },
                    }, drawData.DrawingType);
                    view.SetSketchDimensionValue("D1@Sketch447", W_upperTol);
                    view.SetSketchDimensionValue("D2@Sketch447", W_lowerTol);
                    view.SetSketchDimensionValue("D3@Sketch447", b);
                }

                view = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

           /* // 🔁 Align top view relative to side view AFTER all views are created
            if (sideView != null && topView != null)
            {
                topView.AlignTopViewNextToSideView(
                    sideView.GetRawView(),
                    topView.GetRawView(),
                    offsetMm: 10.0
                );
            }
            else
            {
                Logger.Error("Top view and/or side view were not found for alignment.");
            }*/
        }


        private static void FinalizeDrawing(SldWorks swApp, IDrawingService drawingService, IPartService partService, string outputPdfPath, string outputTiffPath, WedgeData wedgedata, INoteService noteService)
        {
            var model = drawingService.GetModel();
            var conf = new TiffExportSettings(swApp);
            model.GraphicsRedraw2();
            drawingService.ZoomToSheet();
            double overlay_calibration = (double.Parse(wedgedata.OverlayCalibration) / 25400) * wedgedata.OverlayScaling;
            drawingService.DrawCenteredSquareOnSheet(overlay_calibration);
            noteService.InsertOverlayCalibrationNote(wedgedata.OverlayCalibration, overlay_calibration);
            drawingService.SaveAsTiff(outputTiffPath);
            drawingService.SaveAsPdf(outputPdfPath);
            Thread.Sleep(300);

            // Build new resized tiff file name
            string resizedTiff = Path.Combine(
                Path.GetDirectoryName(outputTiffPath),
                Path.GetFileNameWithoutExtension(outputTiffPath) + "_1280_1024.tif"
            );

            conf.ResizeImageSharpHighQuality(outputTiffPath, resizedTiff);
            drawingService.Unlock();
            drawingService.SaveAndCloseDrawing();

            partService.Save(close: true);
        }

    }
}
