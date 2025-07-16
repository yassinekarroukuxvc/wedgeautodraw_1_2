using ClosedXML.Excel;
using SolidWorks.Interop.sldworks;
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

            var noteService = new NoteService(swApp, drawingService.GetModel());

            InsertOverlayNotes(noteService, drawingData, wedgeData);

            FinalizeDrawing(swApp, drawingService, partService, outputPdfPath, outputTiffPath, wedgeData, noteService);

            Logger.Success("Overlay drawing automation completed.");
        }

        private static DrawingService InitializeDrawing(SldWorks swApp, string partPath, string drawingPath, string modPartPath, string modDrawingPath)
        {
            var drawingService = new DrawingService(swApp);
            drawingService.ReplaceReferencedModel(modDrawingPath, partPath, modPartPath);
            drawingService.OpenDrawing(modDrawingPath);
            drawingService.Unlock();
            swApp.Visible = true;
            drawingService.Lock();
            drawingService.Rebuild();

            Logger.Info("Drawing initialized and rebuilt.");
            return drawingService;
        }

        private static void UpdateViewScalesAndPositions(SldWorks swApp, IDrawingService drawingService, DrawingData drawData, WedgeData wedgeData)
        {
            var model = drawingService.GetModel();

            string wedgeId = wedgeData.Metadata.ContainsKey("drawing_number")
                ? wedgeData.Metadata["drawing_number"].ToString()
                : $"Wedge_{Guid.NewGuid()}";

            if (ShouldHideLayer8(wedgeId))
            {
                TryHideLayer8(model);
            }

            string[] viewNames = new[]
            {
                Constants.OverlaySideView,
                Constants.OverlayTopView,
                Constants.OverlayDetailView,
                Constants.OverlaySectionView,
                Constants.OverlaySideView2
            };

            double topSideScale = drawData.ViewScales[Constants.SideView].GetValue(Unit.Millimeter);
            double tl = wedgeData.Dimensions["TL"].GetValue(Unit.Meter);

            foreach (var viewName in viewNames)
            {
                Logger.Info($"Processing view: {viewName}");
                var viewFactory = new StandardViewFactory(model);
                var view = viewFactory.CreateView(viewName);

                var dimensionKeysForView = GetDimensionKeysForView(wedgeData.WedgeType, viewName);

                if (viewName == Constants.OverlaySideView2)
                {
                    view.DeleteOverlaySideView2(wedgeData.Dimensions);
                    view.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, dimensionKeysForView, drawData.DrawingType);
                }

                if (viewName == Constants.OverlayDetailView || viewName == Constants.OverlaySectionView)
                {
                    view.SetViewScale(wedgeData.OverlayScaling);
                }

                if (viewName == Constants.OverlaySideView)
                {
                    view.SetViewScale(topSideScale);
                    view.AlignViewHorizontally(false, tlInMeters: tl);
                    view.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, dimensionKeysForView, drawData.DrawingType);

                    view.SetSketchDimensionValue("D1@Sketch33", 0.19 / topSideScale);
                }

                if (viewName == Constants.OverlayTopView)
                {
                    view.SetViewScale(drawData.ViewScales[Constants.TopView].GetValue(Unit.Millimeter));
                    view.SetViewPosition(drawData.ViewPositions[Constants.TopView]);
                    view.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, dimensionKeysForView, drawData.DrawingType);

                    view.SetSketchDimensionValue("D2@Sketch33", 0.006 / topSideScale);
                    view.SetSketchDimensionValue("D3@Sketch33", 0.015 / topSideScale);

                    double d3 = view.GetSketchDimensionValue("D3@Sketch33");
                    view.SetSketchDimensionValue("D4@Sketch33", d3 / 2);
                }

                if (viewName == Constants.OverlaySectionView)
                {
                    ApplyOverlaySectionViewSettings(view, wedgeData, drawData);
                }

                if (viewName == Constants.OverlayDetailView)
                {
                    ApplyOverlayDetailViewSettings(view, wedgeData, drawData);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static Dictionary<string, string> GetDimensionKeysForView(WedgeType wedgeType, string viewName)
        {
            var keys = new Dictionary<string, string>();

            if (wedgeType == WedgeType.CKVD)
            {
                if (viewName == Constants.OverlaySideView)
                {
                    keys = new()
                    {
                        { "FX", "SelectByName" },
                        { "D3", "SelectByName" },
                        { "FA", "SelectByName" },
                        { "BA", "SelectByName" },
                        { "E", "SelectByName" },
                        { "X", "SelectByName" },
                    };
                }
                else if (viewName == Constants.OverlayTopView)
                {
                    keys = new()
                    {
                        { "TDF", "SelectByName" },
                    };
                }
                else if (viewName == Constants.OverlaySideView2)
                {
                    keys = new()
                    {
                        { "VR", "SelectByName" },
                        { "VW", "SelectByName" },
                    };
                }
                else if (viewName == Constants.OverlayDetailView)
                {
                    keys = new()
                    {
                        { "ISA", "SelectByName" },
                        { "GA", "SelectByName" },
                    };
                }
            }
            else if (wedgeType == WedgeType.COB)
            {
                if (viewName == Constants.OverlaySideView)
                {
                    keys = new()
                    {
                        { "ISA", "SelectByName" },
                        { "BA", "SelectByName" },
                        { "RA", "SelectByName" },
                        { "FD", "SelectByName" },
                        { "E", "SelectByName" },
                        { "X", "SelectByName" },
                    };
                }
                else if (viewName == Constants.OverlayTopView)
                {
                    keys = new()
                    {
                        { "TDF", "SelectByName" },
                        { "TD", "SelectByName" },
                    };
                }
                else if (viewName == Constants.OverlayDetailView)
                {
                    keys = new()
                    {
                        { "ISA", "SelectByName" },
                        { "GA", "SelectByName" },
                        { "FRO", "SelectByName" },
                    };
                }
                else if (viewName == Constants.OverlaySideView2)
                {
                    keys = new()
                    {
                        { "VR", "SelectByName" },
                        { "VW", "SelectByName" },
                    };
                }
            }

            return keys;
        }

        private static void ApplyOverlaySectionViewSettings(IViewService view, WedgeData wedgeData, DrawingData drawData)
        {
            double FL_lowerTol = wedgeData.Dimensions["FL"].GetTolerance(Unit.Meter, "-");
            double FL_upperTol = wedgeData.Dimensions["FL"].GetTolerance(Unit.Meter, "+");

            view.SetOverlayBreaklinePosition(wedgeData.Dimensions, drawData);
            view.CenterViewVertically();
            view.SetBreakLineGap(0.000025);
            view.AlignViewHorizontally(false, tlInMeters: 0);
            view.CenterSectionViewVisuallyVertically(wedgeData.Dimensions);

            view.SetSketchDimensionValue("D1@Sketch474", FL_upperTol);
            view.SetSketchDimensionValue("D3@Sketch474", FL_lowerTol);
        }

        private static void ApplyOverlayDetailViewSettings(IViewService view, WedgeData wedgeData, DrawingData drawData)
        {
            double W_lowerTol = wedgeData.Dimensions["W"].GetTolerance(Unit.Meter, "-");
            double W_upperTol = wedgeData.Dimensions["W"].GetTolerance(Unit.Meter, "+");
            double b = wedgeData.Dimensions["B"].GetTolerance(Unit.Meter, "+");

            view.SetOverlayBreaklinePosition(wedgeData.Dimensions, drawData);
            view.CenterViewVertically();
            view.SetBreakLineGap(0.000025);
            view.AlignViewHorizontally(true, tlInMeters: 0);

            var dimensionKeysForDetailView = GetDimensionKeysForView(wedgeData.WedgeType, Constants.OverlayDetailView);
            view.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, dimensionKeysForDetailView, drawData.DrawingType);

            view.SetSketchDimensionValue("D1@Sketch447", W_upperTol);
            view.SetSketchDimensionValue("D2@Sketch447", W_lowerTol);
            view.SetSketchDimensionValue("D3@Sketch447", b);
        }

        private static void InsertOverlayNotes(INoteService noteService, DrawingData drawingData, WedgeData wedgeData)
        {
            noteService.InsertDimensionNote(
                drawingData.TablePositions[Constants.DimensionTable],
                drawingData.DimensionKeysInTable,
                "DIMENSIONS:",
                drawingData,
                wedgeData.Dimensions
            );

            noteService.InsertCustomNoteAsTable(wedgeData.Coining, drawingData.TablePositions[Constants.CoiningNote]);
        }

        private static void FinalizeDrawing(SldWorks swApp, IDrawingService drawingService, IPartService partService, string outputPdfPath, string outputTiffPath, WedgeData wedgeData, INoteService noteService)
        {
            var model = drawingService.GetModel();
            var conf = new TiffExportSettings(swApp);

            model.GraphicsRedraw2();
            drawingService.ZoomToSheet();

            double overlayCalibration = (double.Parse(wedgeData.OverlayCalibration) / 25400) * wedgeData.OverlayScaling;
            drawingService.DrawCenteredSquareOnSheet(overlayCalibration);
            noteService.InsertOverlayCalibrationNote(wedgeData.OverlayCalibration, overlayCalibration);

            drawingService.SaveAsTiff(outputTiffPath);
            drawingService.SaveAsPdf(outputPdfPath);

            Thread.Sleep(300);

            string resizedTiff = Path.Combine(
                Path.GetDirectoryName(outputTiffPath),
                Path.GetFileNameWithoutExtension(outputTiffPath) + "_1280_1024.tif"
            );

            conf.ResizeImageSharpHighQuality(outputTiffPath, resizedTiff);

            drawingService.Unlock();
            drawingService.SaveAndCloseDrawing();

            partService.Save(close: true);
        }

        private static bool ShouldHideLayer8(string wedgeId)
        {
            var specialWedgeIds = new HashSet<string> { "2026582", "2026989", "2030604" };
            return specialWedgeIds.Contains(wedgeId);
        }

        private static void TryHideLayer8(ModelDoc2 model)
        {
            Logger.Info("Hiding Layer 8 if present.");
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
    }
}
