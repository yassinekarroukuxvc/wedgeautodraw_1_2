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
        Logger.Info("=== Starting Part Automation ===");

        var partService = new PartService(swApp);
        partService.OpenPart(modPartPath);
        partService.ApplyTolerances(wedgeData.Dimensions);
        partService.UpdateEquations(modEquationPath);
        partService.SetEngravedText(wedgeData.EngravedText);
        partService.Rebuild();
        partService.Save();

        Logger.Success("Part automation completed.");
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
        Logger.Info("=== Starting Drawing Automation ===");

        var drawingService = InitializeDrawing(swApp, partPath, drawingPath, modPartPath, modDrawingPath, drawingData);
        CreateAllStandardViews(drawingService, drawingData, wedgeData);
        string sectionViewName = CreateSectionView(drawingService, drawingData, wedgeData);
        FinalizeDrawing(drawingService, sectionViewName, drawingData, wedgeData, partService, swApp, outputPdfPath);

        Logger.Success("Drawing automation completed.");
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

        Logger.Info("Drawing initialized and rebuilt.");
        return drawingService;
    }

    private static void CreateAllStandardViews(IDrawingService drawingService, DrawingData drawData, WedgeData wedgeData)
    {
        var model = drawingService.GetModel();
        CreateFrontView(model, drawData, wedgeData);
        CreateSideView(model, drawData, wedgeData);
        CreateTopView(model, drawData, wedgeData);
        CreateDetailView(model, drawData, wedgeData);

        Logger.Info("Standard views created (Front, Side, Top, Detail).");
    }

    private static string CreateSectionView(IDrawingService drawingService, DrawingData drawingData, WedgeData wedgeData)
    {
        Logger.Info("=== Creating Section View ===");

        var model = drawingService.GetModel();
        var sketchSegment = CreateCuttingLineSketch(drawingService, wedgeData);

        var tempDetailView = new ViewService(Constants.DetailView, ref model);
        var sectionViewName = tempDetailView.CreateSectionView(
            tempDetailView,
            drawingData.ViewPositions[Constants.SectionView],
            sketchSegment,
            wedgeData.Dimensions,
            drawingData);
        drawingService.Rebuild();

        return sectionViewName;
    }

    private static void FinalizeDrawing(
        IDrawingService drawingService,
        string sectionViewName,
        DrawingData drawingData,
        WedgeData wedgeData,
        IPartService partService,
        SldWorks swApp,
        string outputPdfPath)
    {
        Logger.Info("=== Finalizing Drawing ===");

        partService.ToggleSketchVisibility(Constants.SketchEngraving, false);
        partService.ToggleSketchVisibility(Constants.SketchGrooveDimensions, true);
        partService.Save();

        var model = drawingService.GetModel();
        drawingService.Rebuild();

        var sectionView = new ViewService(sectionViewName, ref model);
        sectionView.ReactivateView(ref model);
        sectionView.SetViewScale(drawingData.ViewScales[Constants.SectionView].GetValue(Unit.Millimeter));
        sectionView.InsertModelDimensioning();
        AdjustSectionOrDetailViewDimensions(sectionView, drawingData, wedgeData, Constants.SectionView, new()
        {
            { "F", "SelectByName" },
            { "FL", "SelectByName" },
            { "FR", "SelectByName" },
            { "BR", "SelectByName" }
        });

        partService.ToggleSketchVisibility(Constants.SketchGrooveDimensions, false);
        partService.Save(close: true);
        partService.Unlock();

        CreateDrawingTables(swApp, drawingService, drawingData, wedgeData);
        drawingService.GetModel().GraphicsRedraw2();
        drawingService.SaveAsPdf(outputPdfPath);
        drawingService.Unlock();
        drawingService.SaveAndCloseDrawing();
    }

    private static void CreateDrawingTables(SldWorks swApp, IDrawingService drawingService, DrawingData drawingData, WedgeData wedgeData)
    {
        Logger.Info("=== Creating Tables ===");

        var tableService = new TableService(swApp, drawingService.GetModel());

        tableService.CreateDimensionTable(drawingData.TablePositions[Constants.DimensionTable], drawingData.DimensionKeysInTable, "DIMENSIONS:", drawingData, wedgeData.Dimensions);
        tableService.CreateHowToOrderTable(drawingData.TablePositions[Constants.HowToOrderTable], "HOW TO ORDER", drawingData);

        Logger.Success("Tables created successfully.");
    }

    private static SketchSegment CreateCuttingLineSketch(IDrawingService drawingService, WedgeData wedgeData)
    {
        var model = drawingService.GetModel();
        model.ClearSelection2(true);

        Logger.Info("Creating cutting line sketch for section view.");

        return model.SketchManager.CreateLine(
            0.0,
            -wedgeData.Dimensions["TL"].GetValue(Unit.Millimeter) / 2,
            0.0,
            0.0,
            wedgeData.Dimensions["TL"].GetValue(Unit.Millimeter) / 2,
            0.0
        );
    }

    private static void CreateFrontView(ModelDoc2 model, DrawingData drawData, WedgeData wedgeData)
    {
        var front = new ViewService(Constants.FrontView, ref model);
        front.SetViewScale(drawData.ViewScales[Constants.FrontView].GetValue(Unit.Millimeter));
        front.SetViewPosition(drawData.ViewPositions[Constants.FrontView]);
        front.SetBreaklinePosition(wedgeData.Dimensions, drawData);
        front.CreateFixedCenterline(wedgeData.Dimensions, drawData);
        front.SetBreakLineGap(drawData.BreaklineData["Front_viewBreaklineGap"].GetValue(Unit.Meter));
        front.SetPositionAndNameDimensioning(wedgeData.Dimensions, drawData.DimensionStyles, new()
        {
            { "TL", "SelectByName" },
            { "EngravingStart", "SelectByName" }
        });
    }

    private static void CreateSideView(ModelDoc2 model, DrawingData drawData, WedgeData wedgeData)
    {
        var side = new ViewService(Constants.SideView, ref model);
        side.SetViewPosition(drawData.ViewPositions[Constants.SideView]);
        side.SetBreaklinePosition(wedgeData.Dimensions, drawData);
        side.SetBreakLineGap(drawData.BreaklineData["Side_viewBreaklineGap"].GetValue(Unit.Meter));
        side.CreateFixedCenterline(wedgeData.Dimensions, drawData);
        side.SetPositionAndNameDimensioning(wedgeData.Dimensions, drawData.DimensionStyles, new()
        {
            { "FA", "SelectByName" },
            { "BA", "SelectByName" },
            { "E", "SelectByName" },
            { "FX", "SelectByName" }
        });
    }

    private static void CreateTopView(ModelDoc2 model, DrawingData drawData, WedgeData wedgeData)
    {
        var top = new ViewService(Constants.TopView, ref model);
        top.SetViewPosition(drawData.ViewPositions[Constants.TopView]);
        top.CreateFixedCentermark(wedgeData.Dimensions, drawData);
        top.SetPositionAndNameDimensioning(wedgeData.Dimensions, drawData.DimensionStyles, new()
        {
            { "TDF", "SelectByName" },
            { "TD", "SelectByName" }
        });
        top.SetPositionAndLabelDatumFeature(wedgeData.Dimensions, drawData.DimensionStyles, "A");
    }

    private static void CreateDetailView(ModelDoc2 model, DrawingData drawData, WedgeData wedgeData)
    {
        var detail = new ViewService(Constants.DetailView, ref model);
        detail.SetViewScale(drawData.ViewScales[Constants.DetailView].GetValue(Unit.Millimeter));
        detail.SetViewPosition(drawData.ViewPositions[Constants.DetailView]);
        detail.SetBreaklinePosition(wedgeData.Dimensions, drawData);
        detail.SetBreakLineGap(drawData.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Meter));
        detail.CreateFixedCenterline(wedgeData.Dimensions, drawData);
        detail.SetPositionAndValuesAndLabelGeometricTolerance(wedgeData.Dimensions, drawData.DimensionStyles, Constants.DatumFeatureLabel);
        AdjustSectionOrDetailViewDimensions(detail, drawData, wedgeData, Constants.DetailView, new()
        {
            { "ISA", "SelectByName" },
            { "GA", "SelectByName" },
            { "B", "SelectByName" },
            { "W", "SelectByName" },
            { "GD", "SelectByName" },
            { "GR", "SelectByName" }
        });
    }

    private static void AdjustSectionOrDetailViewDimensions(ViewService view, DrawingData drawData, WedgeData wedgeData, string viewKey, Dictionary<string, string> dimensionTypes)
    {
        var defaultPositions = view.GetDefaultModelDimensionPositions();
        var scale = drawData.ViewScales[viewKey].GetValue(Unit.Millimeter);
        var FL = wedgeData.Dimensions.ContainsKey("FL") ? wedgeData.Dimensions["FL"].GetValue(Unit.Millimeter) : 0;
        var GD = wedgeData.Dimensions.ContainsKey("GD") ? wedgeData.Dimensions["GD"].GetValue(Unit.Millimeter) : 0;

        foreach (var kvp in defaultPositions)
        {
            string dimName = kvp.Key;
            double[] pos = kvp.Value;
            double[] adjustedPos = SectionDetailViewsAdjuster.ApplyOffset(dimName, pos, scale, FL, GD);

            if (drawData.DimensionStyles.ContainsKey(dimName))
                drawData.DimensionStyles[dimName].Position = new DataStorage(adjustedPos);
            else
                drawData.DimensionStyles[dimName] = new DimensioningStorage(new DataStorage(adjustedPos));
        }

        view.SetPositionAndNameDimensioning(wedgeData.Dimensions, drawData.DimensionStyles, dimensionTypes);
    }
}