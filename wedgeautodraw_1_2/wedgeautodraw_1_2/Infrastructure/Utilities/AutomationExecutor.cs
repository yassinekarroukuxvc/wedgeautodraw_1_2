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
        //EquationFileUpdater.EnsureAllEquationsExist(partService.GetModel(), wedgeData);
        //partService.SetEngravedText(wedgeData.EngravedText);
        partService.SetEngravedText("Test");
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
        sectionView.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawingData.DimensionStyles, new()
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
        front.CreateCenterline(wedgeData.Dimensions, drawData);
        front.SetBreakLineGap(drawData.BreaklineData["Front_viewBreaklineGap"].GetValue(Unit.Meter));
        front.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
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
        side.CreateCenterline(wedgeData.Dimensions, drawData);
        side.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
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
        top.CreateCentermark(wedgeData.Dimensions, drawData);
        top.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
        {
            { "TDF", "SelectByName" },
            { "TD", "SelectByName" }
        });
        top.PlaceDatumFeatureLabel(wedgeData.Dimensions, drawData.DimensionStyles, "A");
    }

    private static void CreateDetailView(ModelDoc2 model, DrawingData drawData, WedgeData wedgeData)
    {
        var detail = new ViewService(Constants.DetailView, ref model);
        detail.SetViewScale(drawData.ViewScales[Constants.DetailView].GetValue(Unit.Millimeter));
        detail.SetViewPosition(drawData.ViewPositions[Constants.DetailView]);
        detail.SetBreaklinePosition(wedgeData.Dimensions, drawData);
        detail.SetBreakLineGap(drawData.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Meter));
        detail.CreateCenterline(wedgeData.Dimensions, drawData);
        detail.PlaceGeometricToleranceFrame(wedgeData.Dimensions, drawData.DimensionStyles, Constants.DatumFeatureLabel);
        detail.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
        {
            { "ISA", "SelectByName" },
            { "GA", "SelectByName" },
            { "B", "SelectByName" },
            { "W", "SelectByName" },
            { "GD", "SelectByName" },
            { "GR", "SelectByName" }
        });
    }
    private static ViewService EnsureOrCreateView(string viewName, ModelDoc2 model, string partPath)
    {
        var drawingDoc = model as DrawingDoc;
        if (drawingDoc == null)
        {
            Logger.Error("Model is not a DrawingDoc.");
            return null;
        }

        object[] views = drawingDoc.GetViews() as object[];
        if (views != null)
        {
            foreach (object obj in views)
            {
                if (obj is View view)
                {
                    string lowerName = view.Name?.ToLowerInvariant() ?? "";
                    string lowerTarget = viewName.ToLowerInvariant();

                    // Fuzzy match: view name contains "front", "side", etc.
                    if (lowerName.Contains(lowerTarget.Replace("_view", "")))
                    {
                        Logger.Info($"Matched existing view '{view.Name}' to '{viewName}'");
                        return new ViewService(view.Name, ref model);
                    }
                }
            }
        }

        Logger.Info($"View '{viewName}' not found. Creating...");

        string orientation = viewName switch
        {
            "Front_view" => "*Front",
            "Top_view" => "*Top",
            "Side_view" => "*Right",
            _ => "*Front"
        };

        View createdView = drawingDoc.CreateDrawViewFromModelView3(partPath, orientation, 0.0, 0.0, 0.0);
        if (createdView == null)
        {
            Logger.Error($"Failed to create view: {viewName} with orientation '{orientation}'");
            return null;
        }

        Logger.Success($"Successfully created view '{viewName}' from orientation '{orientation}'.");

        return new ViewService(createdView.Name, ref model);
    }


}