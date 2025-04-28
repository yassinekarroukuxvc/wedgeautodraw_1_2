using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Services;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Utilities;

public static class AutomationExecutor
{
    public static IPartService RunPartAutomation(SldWorks swApp, string modEquationPath, string modPartPath, WedgeData wedgeData)
    {
        var partService = new PartService(swApp);
        partService.OpenPart(modPartPath);
        partService.ApplyTolerances(wedgeData.Dimensions);
        partService.UpdateEquations(modEquationPath);
        partService.SetEngravedText(wedgeData.EngravedText);
        partService.Rebuild();
        partService.Save();
        return partService;
    }

    public static void RunDrawingAutomation(
        SldWorks swApp,
        IPartService partService,
        string partPath,
        string drawingPath,
        string modPartPath,
        string modDrawingPath,
        string modEquationPath,
        DrawingData drawingData,
        WedgeData wedgeData,
        string outputPdfPath)
    {
        var drawingService = InitializeDrawing(swApp, partPath, drawingPath, modPartPath, modDrawingPath, drawingData);

        CreateStandardViews(drawingService, drawingData, wedgeData);

        var sectionViewName = CreateAndConfigureSectionView(drawingService, drawingData, wedgeData);

        FinalizeDrawing(swApp,drawingService, sectionViewName, drawingData, wedgeData, partService, outputPdfPath);
    }

    private static DrawingService InitializeDrawing(
        SldWorks swApp,
        string partPath,
        string drawingPath,
        string modPartPath,
        string modDrawingPath,
        DrawingData drawingData)
    {
        var drawingService = new DrawingService(swApp);
        drawingService.ReplaceReferencedModel(modDrawingPath, partPath, modPartPath);
        drawingService.OpenDrawing(modDrawingPath);
        drawingService.SetSummaryInformation(drawingData);
        drawingService.SetCustomProperties(drawingData);
        drawingService.Rebuild();
        return drawingService;
    }

    private static void CreateStandardViews(IDrawingService drawingService, DrawingData draw, WedgeData wed)
    {
        var model = drawingService.GetModel();
        CreateFrontView(model, draw, wed);
        CreateSideView(model, draw, wed);
        CreateTopView(model, draw, wed);
        CreateDetailView(model, draw, wed);
    }

    private static string CreateAndConfigureSectionView(
        IDrawingService drawingService,
        DrawingData drawingData,
        WedgeData wedgeData)
    {
        var model = drawingService.GetModel();

        // Create cutting line sketch
        var sketchSegment = CreateCuttingLineSketch(drawingService, wedgeData);

        // Create section view
        var tempDetailView = new ViewService("Detail_view", ref model);
        string actualSectionViewName = tempDetailView.CreateSectionView(
            new ViewService("Detail_view", ref model),
            drawingData.ViewPositions["Section_view"],
            sketchSegment,
            wedgeData.Dimensions,
            drawingData
        );

        drawingService.SaveDrawing();
        drawingService.SaveAndCloseDrawing();
        return actualSectionViewName;
    }

    private static void FinalizeDrawing(
        SldWorks swApp,
        IDrawingService drawingService,
        string sectionViewName,
        DrawingData drawingData,
        WedgeData wedgeData,
        IPartService partService,
        string outputPdfPath)
    {
        partService.ToggleSketchVisibility("sketch_engraving", false);
        partService.ToggleSketchVisibility("sketch_groove_dimensions", true);
        partService.Save();

        drawingService.Reopen();
        var model = drawingService.GetModel();
        drawingService.Rebuild();

        var sectionView = new ViewService(sectionViewName, ref model);
        sectionView.ReactivateView(ref model);
        sectionView.SetViewScale(drawingData.ViewScales["Section_view"].GetValue(Unit.Millimeter));
        sectionView.InsertModelDimensioning();

        AdjustSectionViewDimensions(sectionView, drawingData, wedgeData);
    
        partService.ToggleSketchVisibility("sketch_groove_dimensions", false);
        partService.Save(close: true);
        partService.Unlock();

        // Create tables
        var tableService = new TableService(swApp, drawingService.GetModel());
        tableService.CreateDimensionTable(drawingData.TablePositions["dimension"], drawingData.DimensionKeysInTable, "DIMENSIONS:", drawingData, wedgeData.Dimensions);
        //tableService.CreateLabelAsTable(drawingData.TablePositions["label_as"], drawingData);
        //tableService.CreatePolishTable(drawingData.TablePositions["polish"], drawingData);
        tableService.CreateHowToOrderTable(drawingData.TablePositions["how_to_order"], "HOW TO ORDER", drawingData);

        drawingService.ZoomToFit();
        drawingService.SaveAsPdf(outputPdfPath);
        drawingService.Unlock();
        drawingService.SaveAndCloseDrawing();
    }

    private static void AdjustSectionViewDimensions(ViewService sectionView, DrawingData drawingData, WedgeData wedgeData)
    {
        var defaultPositions = sectionView.GetDefaultModelDimensionPositions();
        var secv = drawingData.ViewScales["Section_view"].GetValue(Unit.Millimeter);
        var FL = wedgeData.Dimensions["FL"].GetValue(Unit.Millimeter);
        var GD = wedgeData.Dimensions["GD"].GetValue(Unit.Millimeter);

        foreach (var kvp in defaultPositions)
        {
            string dimName = kvp.Key;
            double[] pos = kvp.Value;

            double[] adjustedPos = SectionViewAdjuster.ApplyOffset(dimName, pos, secv, FL, GD);

            if (drawingData.DimensionStyles.ContainsKey(dimName))
            {
                drawingData.DimensionStyles[dimName].Position = new DataStorage(adjustedPos);
            }
            else
            {
                drawingData.DimensionStyles[dimName] = new DimensioningStorage(new DataStorage(adjustedPos));
            }
        }

        sectionView.SetPositionAndNameDimensioning(wedgeData.Dimensions, drawingData.DimensionStyles, new()
        {
            { "F", "SelectByName" },
            { "FL", "SelectByName" },
            { "FR", "SelectByName" },
            { "BR", "SelectByName" }
        });
    }

    private static SketchSegment CreateCuttingLineSketch(IDrawingService drawingService, WedgeData wedgeData)
    {
        var model = drawingService.GetModel();
        model.ClearSelection2(true);

        return model.SketchManager.CreateLine(
            0.0,
            -wedgeData.Dimensions["TL"].GetValue(Unit.Millimeter) / 2,
            0.0,
            0.0,
            wedgeData.Dimensions["TL"].GetValue(Unit.Millimeter) / 2,
            0.0
        );
    }

    private static void CreateFrontView(ModelDoc2 model, DrawingData draw, WedgeData wed)
    {
        var front = new ViewService("Front_view", ref model);
        front.SetViewScale(draw.ViewScales["Front_view"].GetValue(Unit.Millimeter));
        front.SetViewPosition(draw.ViewPositions["Front_view"]);
        front.SetBreaklinePosition(wed.Dimensions, draw);
        front.CreateFixedCenterline(wed.Dimensions, draw);
        front.SetBreakLineGap(draw.BreaklineData["Front_viewBreaklineGap"].GetValue(Unit.Meter));
        front.SetPositionAndNameDimensioning(wed.Dimensions, draw.DimensionStyles, new()
        {
            { "TL", "SelectByName" },
            { "EngravingStart", "SelectByName" }
        });
    }

    private static void CreateSideView(ModelDoc2 model, DrawingData draw, WedgeData wed)
    {
        var side = new ViewService("Side_view", ref model);
        side.SetViewPosition(draw.ViewPositions["Side_view"]);
        side.SetBreaklinePosition(wed.Dimensions, draw);
        side.SetBreakLineGap(draw.BreaklineData["Side_viewBreaklineGap"].GetValue(Unit.Meter));
        side.CreateFixedCenterline(wed.Dimensions, draw);
        side.SetPositionAndNameDimensioning(wed.Dimensions, draw.DimensionStyles, new()
        {
            { "FA", "SelectByName" },
            { "BA", "SelectByName" },
            { "E", "SelectByName" },
            { "FX", "SelectByName" }
        });
    }

    private static void CreateTopView(ModelDoc2 model, DrawingData draw, WedgeData wed)
    {
        var top = new ViewService("Top_view", ref model);
        top.SetViewPosition(draw.ViewPositions["Top_view"]);
        top.CreateFixedCentermark(wed.Dimensions, draw);
        top.SetPositionAndNameDimensioning(wed.Dimensions, draw.DimensionStyles, new()
        {
            { "TDF", "SelectByName" },
            { "TD", "SelectByName" }
        });
        top.SetPositionAndLabelDatumFeature(wed.Dimensions, draw.DimensionStyles, "A");
    }

    private static void CreateDetailView(ModelDoc2 model, DrawingData draw, WedgeData wed)
    {
        var detail = new ViewService("Detail_view", ref model);
        detail.SetViewScale(draw.ViewScales["Detail_view"].GetValue(Unit.Millimeter));
        detail.SetViewPosition(draw.ViewPositions["Detail_view"]);
        detail.SetBreaklinePosition(wed.Dimensions, draw);
        detail.SetBreakLineGap(draw.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Meter));
        detail.CreateFixedCenterline(wed.Dimensions, draw);
        detail.SetPositionAndValuesAndLabelGeometricTolerance(wed.Dimensions, draw.DimensionStyles, "A");
        detail.SetPositionAndNameDimensioning(wed.Dimensions, draw.DimensionStyles, new()
        {
                     /*{"ISA", "SelectByName"},
                     {"GA" , "SelectByName"},
                     {"B"  , "SelectByName"},
                     {"W"  , "SelectByName"},
                     {"GD" , "SelectByName"},
                     {"GR" , "SelectByName"}*/
        });
    }
}
