using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Factories;
using wedgeautodraw_1_2.Infrastructure.Helpers;
using wedgeautodraw_1_2.Infrastructure.Services;

namespace wedgeautodraw_1_2.Infrastructure.Executors;

public class ProductionDrawingAutomationExecutor : IDrawingAutomationExecutor
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
        sectionView.InsertModelDimensioning(drawingData.DrawingType);
        sectionView.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawingData.DimensionStyles, new()
        {
            { "F", "SelectByName" },
            { "FL", "SelectByName" },
            { "FR", "SelectByName" },
            { "BR", "SelectByName" }
        },drawingData.DrawingType);

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
        double scale = drawData.ViewScales[Constants.FrontView].GetValue(Unit.Millimeter);

        var viewFactory = new StandardViewFactory(model);
        var front = viewFactory.CreateView(Constants.FrontView);
        front.SetViewPosition(drawData.ViewPositions[Constants.FrontView]);
        //front.SetBreaklinePosition(wedgeData.Dimensions, drawData);
        front.SetFrontSideViewBreakline(wedgeData.Dimensions);
        front.CreateCenterline(wedgeData.Dimensions, drawData);
        front.SetBreakLineGap(drawData.BreaklineData["Front_viewBreaklineGap"].GetValue(Unit.Meter));
        drawData.ViewScales[Constants.FrontView] = new DataStorage(scale);
        front.SetViewScale(drawData.ViewScales[Constants.FrontView].GetValue(Unit.Millimeter));
        front.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
            {
                { "TL", "SelectByName" },
                { "EngravingStart", "SelectByName" },
                { "D2", "SelectByName" },
                { "VW", "SelectByName" },
            },drawData.DrawingType);
        front.DeleteAnnotationsByName(new[] { "GR" });
     
    }

    private static void CreateSideView(ModelDoc2 model, DrawingData drawData, WedgeData wedgeData)
    {
        var viewFactory = new StandardViewFactory(model);
        var side = viewFactory.CreateView(Constants.SideView);
        side.SetViewPosition(drawData.ViewPositions[Constants.SideView]);
        //side.SetBreaklinePosition(wedgeData.Dimensions, drawData);
        side.SetFrontSideViewBreakline(wedgeData.Dimensions);
        side.SetBreakLineGap(drawData.BreaklineData["Side_viewBreaklineGap"].GetValue(Unit.Meter));
        side.CreateCenterline(wedgeData.Dimensions, drawData);

        var dimensionKeys = new Dictionary<string, string>
        {
            { "FA", "SelectByName" },
            { "BA", "SelectByName" },
            /*{ "E", "SelectByName" },*/
            { "FX", "SelectByName" }
        };

        side.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, dimensionKeys,drawData.DrawingType);
    }


    private static void CreateTopView(ModelDoc2 model, DrawingData drawData, WedgeData wedgeData)
    {
        var viewFactory = new StandardViewFactory(model);
        var top = viewFactory.CreateView(Constants.TopView);
        top.SetViewPosition(drawData.ViewPositions[Constants.TopView]);
        top.CreateCentermark(wedgeData.Dimensions, drawData);
        top.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
        {
            { "TDF", "SelectByName" },
            { "TD", "SelectByName"},
            { "DatumFeature", "SelectByName"}
        }, drawData.DrawingType);
        top.PlaceDatumFeatureLabel(wedgeData.Dimensions, drawData.DimensionStyles, "A");
    }

    private static void CreateDetailView(ModelDoc2 model, DrawingData drawData, WedgeData wedgeData)
    {
        var viewFactory = new StandardViewFactory(model);
        var detail = viewFactory.CreateView(Constants.DetailView);
        detail.SetViewScale(drawData.ViewScales[Constants.DetailView].GetValue(Unit.Millimeter));
        detail.SetViewPosition(drawData.ViewPositions[Constants.DetailView]);
        //detail.SetBreaklinePosition(wedgeData.Dimensions, drawData);
        detail.SetBreakLineGap(drawData.BreaklineData["Detail_viewBreaklineGap"].GetValue(Unit.Meter));
        detail.SetDetailViewDynamicBreakline(wedgeData.Dimensions);
        detail.CreateCenterline(wedgeData.Dimensions, drawData);
        detail.PlaceGeometricToleranceFrame(wedgeData.Dimensions, drawData.DimensionStyles, Constants.DatumFeatureLabel);
        detail.InsertModelDimensioning(drawData.DrawingType);
        detail.ApplyDimensionPositionsAndNames(wedgeData.Dimensions, drawData.DimensionStyles, new()
        {
            { "ISA", "SelectByName" },
            { "GA", "SelectByName" },
            { "B", "SelectByName" },
            { "W", "SelectByName" },
            { "GD", "SelectByName" },
            { "GR", "SelectByName" }
        }, drawData.DrawingType);
    }

}

