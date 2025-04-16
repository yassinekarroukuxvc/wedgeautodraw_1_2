using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Services;

namespace wedgeautodraw_1_2.Infrastructure.Utilities;

public static class AutomationExecutor
{
    public static void RunPartAutomation(SldWorks swApp, string partPath, string equationPath, string modPartPath, string modEquationPath, WedgeData wedgeData)
    {
        IPartService partService = new PartService(swApp);
        partService.OpenPart(modPartPath);
        partService.ApplyTolerances(wedgeData.Dimensions);
        partService.UpdateEquations(modEquationPath);
        partService.SetEngravedText(wedgeData.EngravedText);
        partService.Rebuild();
        partService.Save(close: true);
    }

    public static void RunDrawingAutomation(SldWorks swApp, string drawingPath, string modDrawingPath, string partPath, string modPartPath, DrawingData drawingData, WedgeData wedgeData, string outputPdfPath)
    {
        if (File.Exists(modDrawingPath)) File.Delete(modDrawingPath);
        File.Copy(drawingPath, modDrawingPath);

        IDrawingService drawingService = new DrawingService(swApp);
        drawingService.ReplaceReferencedModel(modDrawingPath, partPath, modPartPath);
        drawingService.OpenDrawing(modDrawingPath);
        drawingService.SetSummaryInformation(drawingData);
        drawingService.SetCustomProperties(drawingData);
        drawingService.Rebuild();

        CreateFrontView(drawingService, drawingData, wedgeData);
        CreateSideView(drawingService, drawingData, wedgeData);
        CreateTopView(drawingService, drawingData, wedgeData);

        // Store created detail view to reuse for section view
        var detailView = CreateDetailView(drawingService, drawingData, wedgeData);
        CreateSectionView(drawingService, drawingData, wedgeData, detailView);

        ITableService tableService = new TableService(swApp, drawingService.GetModel());
        tableService.CreateDimensionTable(drawingData.TablePositions["dimension"], drawingData.DimensionKeysInTable, "DIMENSIONS:", drawingData, wedgeData.Dimensions);
        tableService.CreateLabelAsTable(drawingData.TablePositions["label_as"], drawingData);
        tableService.CreatePolishTable(drawingData.TablePositions["polish"], drawingData);
        tableService.CreateHowToOrderTable(drawingData.TablePositions["how_to_order"], "HOW TO ORDER", drawingData);

        drawingService.ZoomToFit();
        drawingService.SaveAsPdf(outputPdfPath);
        drawingService.SaveAndCloseDrawing();
    }

    private static void CreateFrontView(IDrawingService drawingService, DrawingData draw, WedgeData wed)
    {
        ModelDoc2 model = drawingService.GetModel();
        IViewService view = new ViewService("Front_view", ref model);

        view.SetViewScale(draw.ViewScales["Front_view"].GetValue(Unit.Millimeter));
        view.SetViewPosition(draw.ViewPositions["Front_view"]);
        view.SetBreaklinePosition(wed.Dimensions, draw);
        view.CreateFixedCenterline(wed.Dimensions, draw);
        view.SetBreakLineGap(draw.BreaklineData["Front_viewBreaklineGap"].GetValue(Unit.Meter));

        view.SetPositionAndNameDimensioning(wed.Dimensions, draw.DimensionStyles, new Dictionary<string, string>
        {
            {"TL", "SelectByName"},
            {"EngravingStart", "SelectByName"}
        });
    }

    private static void CreateSideView(IDrawingService drawingService, DrawingData draw, WedgeData wed)
    {
        ModelDoc2 model = drawingService.GetModel();
        IViewService view = new ViewService("Side_view", ref model);

        view.SetViewPosition(draw.ViewPositions["Side_view"]);
        view.SetBreaklinePosition(wed.Dimensions, draw);
        view.SetBreakLineGap(draw.BreaklineData["Side_viewBreaklineGap"].GetValue(Unit.Meter));
        view.CreateFixedCenterline(wed.Dimensions, draw);

        view.SetPositionAndNameDimensioning(wed.Dimensions, draw.DimensionStyles, new Dictionary<string, string>
        {
            {"FA", "SelectByName"},
            {"BA", "SelectByName"},
            {"E",  "SelectByName"},
            {"FX", "SelectByName"}
        });
    }

    private static void CreateTopView(IDrawingService drawingService, DrawingData draw, WedgeData wed)
    {
        ModelDoc2 model = drawingService.GetModel();
        IViewService view = new ViewService("Top_view", ref model);

        view.SetViewPosition(draw.ViewPositions["Top_view"]);
        view.CreateFixedCentermark(wed.Dimensions, draw);

        view.SetPositionAndNameDimensioning(wed.Dimensions, draw.DimensionStyles, new Dictionary<string, string>
        {
            {"TDF", "SelectByName"},
            {"TD",  "SelectByName"}
        });

        view.SetPositionAndLabelDatumFeature(wed.Dimensions, draw.DimensionStyles, "A");
    }

    private static IViewService CreateDetailView(IDrawingService drawingService, DrawingData draw, WedgeData wed)
    {
        ModelDoc2 model = drawingService.GetModel();
        IViewService view = new ViewService("Detail_view", ref model);

        view.SetViewScale(draw.ViewScales["Detail_view"].GetValue(Unit.Millimeter));
        view.SetViewPosition(draw.ViewPositions["Detail_view"]);
        view.SetBreaklinePosition(wed.Dimensions, draw);
        view.SetBreakLineGap(draw.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Meter));
        view.CreateFixedCenterline(wed.Dimensions, draw);

        view.SetPositionAndNameDimensioning(wed.Dimensions, draw.DimensionStyles, new Dictionary<string, string>
        {
            {"ISA", "SelectByName"},
            /*{"GA",  "SelectByName"},
            {"B",   "SelectByName"},
            {"W",   "SelectByName"},
            {"GD",  "SelectByName"},
            {"GR",  "SelectByName"}*/
        });

        view.SetPositionAndValuesAndLabelGeometricTolerance(wed.Dimensions, draw.DimensionStyles, "A");

        return view;
    }

    private static void CreateSectionView(IDrawingService drawingService, DrawingData draw, WedgeData wed, IViewService detailView)
    {
        ModelDoc2 model = drawingService.GetModel();
        model.ForceRebuild3(false);

        try
        {
            Console.WriteLine("🧱 Starting section view creation...");

            IViewService sectionView = new ViewService("Section_view", ref model);
            sectionView.SetViewScale(draw.ViewScales["Section_view"].GetValue(Unit.Millimeter));

            model.ClearSelection2(true);
            var sketchSegment = model.SketchManager.CreateLine(
                0.0,
                -wed.Dimensions["TL"].GetValue(Unit.Inch) / 2,
                0.0,
                0.0,
                wed.Dimensions["TL"].GetValue(Unit.Inch) / 2,
                0.0);

            if (sketchSegment == null)
            {
                Console.WriteLine("❌ Failed to create the sketch segment for the section line.");
                return;
            }

            bool created = sectionView.CreateSectionView(detailView, draw.ViewPositions["Section_view"], sketchSegment, wed.Dimensions, draw);

            if (!created)
            {
                Console.WriteLine("❌ Section view creation failed.");
                return;
            }

            sectionView.ReactivateView(ref model);
            bool dimInserted = sectionView.InsertModelDimensioning();
            if (!dimInserted)
            {
                Console.WriteLine("⚠️ Section view created but no dimensions were inserted.");
            }

            bool positioned = sectionView.SetPositionAndNameDimensioning(wed.Dimensions, draw.DimensionStyles, new Dictionary<string, string>
                {
                    {"F",  "SelectByName"},
                    {"FL", "SelectByName"},
                    {"FR", "SelectByName"},
                    {"BR", "SelectByName"}
                });

            if (!positioned)
            {
                Console.WriteLine("⚠️ Dimensions were not positioned properly in the section view.");
            }

            Console.WriteLine("✅ Section view created and configured.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("❌ Exception during section view creation: " + ex.Message);
        }
    }
}