using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services.ViewServices;

public class SectionViewCreator
{
    private readonly ModelDoc2 _model;

    public SectionViewCreator(ModelDoc2 model)
    {
        _model = model;
    }

    public string Create(DataStorage position, SketchSegment sketchSegment, NamedDimensionValues wedgeDimensions, DrawingData drawData)
    {
        if (_model == null || sketchSegment == null)
        {
            Logger.Warn("Cannot create section view. Model or SketchSegment is null.");
            return null;
        }

        try
        {
            Logger.Info("Starting section view creation...");

            _model.ClearSelection2(true);
            bool selected = sketchSegment.Select4(false, null);

            if (!selected)
            {
                Logger.Warn("Failed to select cutting sketch segment.");
                return null;
            }

            if (_model is not DrawingDoc drawingDoc)
            {
                Logger.Warn("Active model is not a DrawingDoc.");
                return null;
            }

            var view = drawingDoc.CreateSectionViewAt5(
                position.GetValues(Unit.Meter)[0],
                position.GetValues(Unit.Meter)[1],
                0.0,
                "",
                (int)swCreateSectionViewAtOptions_e.swCreateSectionView_ChangeDirection,
                null,
                0.01
            );

            if (view is null)
            {
                Logger.Warn("Section view creation returned null.");
                return null;
            }

            if (view.GetSection() is DrSection swSection)
            {
                swSection.SetAutoHatch(true);
                swSection.SetLabel2(Constants.SectionView);
            }

            Logger.Success($"Section view created successfully: {view.Name}");
            return view.Name;
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception during section view creation: {ex.Message}");
            return null;
        }
    }
}
