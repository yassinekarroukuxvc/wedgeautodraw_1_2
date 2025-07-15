using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using wedgeautodraw_1_2.Infrastructure.Helpers;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Core.Enums;

namespace wedgeautodraw_1_2.Infrastructure.Services.ViewServices
{
    public class ModelViewService
    {
        private readonly ModelDoc2 _modelDoc;
        private readonly DrawingDoc _drawingDoc;

        public ModelViewService(ModelDoc2 modelDoc)
        {
            _modelDoc = modelDoc;
            _drawingDoc = modelDoc as DrawingDoc;
        }

        public bool InsertStandardView(string partPath, string viewname = "*Front", double x = 0.15, double y = 0.15)
        {
            // THE VIEW NAMES CAN BE *Front *Top *Right *Left *Bottom *Back
            if (_drawingDoc == null)
            {
                Logger.Error("Model is not a DrawingDoc.");
                return false;
            }

            try
            {
                View view = _drawingDoc.CreateDrawViewFromModelView3(partPath, viewname, x, y, 0);
                if (view == null)
                {
                    Logger.Warn($"Failed to insert standard view from '{partPath}' with viewname '{viewname}'.");
                    return false;
                }

                Logger.Success($"Inserted standard view at X: {x}, Y: {y}.");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"Exception inserting standard view: {ex.Message}");
                return false;
            }
        }

        public string InsertSectionView(DataStorage position, SketchSegment sketchSegment, string label = "A")
        {
            if (_modelDoc == null || sketchSegment == null)
            {
                Logger.Warn("ModelDoc or SketchSegment is null. Cannot create section view.");
                return null;
            }

            try
            {
                Logger.Info("Starting section view creation...");

                _modelDoc.ClearSelection2(true);
                bool selected = sketchSegment.Select4(false, null);

                if (!selected)
                {
                    Logger.Warn("Failed to select cutting sketch segment.");
                    return null;
                }

                var view = _drawingDoc.CreateSectionViewAt5(
                    position.GetValues(Unit.Meter)[0],
                    position.GetValues(Unit.Meter)[1],
                    0.0,
                    "",
                    (int)swCreateSectionViewAtOptions_e.swCreateSectionView_ChangeDirection,
                    null,
                    0.01
                );

                if (view == null)
                {
                    Logger.Warn("Section view creation returned null.");
                    return null;
                }

                if (view.GetSection() is DrSection swSection)
                {
                    swSection.SetAutoHatch(true);
                    swSection.SetLabel2(label);
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

        public string InsertDetailView(double x, double y, string label = "A")
        {
            if (_drawingDoc == null)
            {
                Logger.Warn("DrawingDoc is null. Cannot create detail view.");
                return null;
            }

            try
            {
                Logger.Info("Starting detail view creation...");

                View detailView = (View)_drawingDoc.CreateDetailViewAt3(
                    x,
                    y,
                    0,     // mode
                    0,     // style
                    1.0,   // scale1
                    1.0,   // scale2
                    label,
                    1,     // foreshortened as int (1 = true)
                    false  // someOption
                );


                if (detailView == null)
                {
                    Logger.Warn("Detail view creation returned null.");
                    return null;
                }

                Logger.Success($"Detail view created successfully: {detailView.GetName2()}");
                return detailView.GetName2();
            }
            catch (Exception ex)
            {
                Logger.Error($"Exception during detail view creation: {ex.Message}");
                return null;
            }
        }
        
        public bool InsertBreakView(string viewName, double firstPos, double secondPos, double gap = 0.01)
        {
            if (_drawingDoc == null)
            {
                Logger.Error("DrawingDoc is null. Cannot insert break view.");
                return false;
            }

            try
            {
                Logger.Info($"Starting break view creation on '{viewName}'...");

                // Activate and select the view
                bool activated = _drawingDoc.ActivateView(viewName);
                if (!activated)
                {
                    Logger.Warn($"Failed to activate view '{viewName}'.");
                    return false;
                }

                bool selected = _modelDoc.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                if (!selected)
                {
                    Logger.Warn($"Failed to select view '{viewName}'.");
                    return false;
                }

                var selectionMgr = (SelectionMgr)_modelDoc.SelectionManager;
                var swView = (View)selectionMgr.GetSelectedObject6(1, -1);

                if (swView == null)
                {
                    Logger.Warn("Failed to get selected view object.");
                    return false;
                }

                // Insert break line (initial line positions approximate)
                BreakLine swBreakLine = (BreakLine)swView.InsertBreak(0, firstPos, secondPos, 1);
                if (swBreakLine == null)
                {
                    Logger.Warn("Failed to insert break lines.");
                    return false;
                }

                // Adjust position of break line
                bool setPos = swBreakLine.SetPosition(firstPos, secondPos);
                if (!setPos)
                {
                    Logger.Warn("Failed to set positions for break lines.");
                }

                _modelDoc.EditRebuild3();

                // Apply the break
                _drawingDoc.BreakView();

                Logger.Success($"Break view inserted and applied successfully on '{viewName}'.");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"Exception during break view creation: {ex.Message}");
                return false;
            }
        }
        public List<string> GetAllViewNames()
        {
            var viewNames = new List<string>();

            if (_drawingDoc == null)
            {
                Logger.Warn("DrawingDoc is null. Cannot get view names.");
                return viewNames;
            }

            try
            {
                // Get the first view (sheet view)
                View sheetView = _drawingDoc.GetFirstView() as View;
                if (sheetView == null)
                {
                    Logger.Warn("No sheet view found.");
                    return viewNames;
                }

                // Start with the first model view after the sheet view
                View view = sheetView.GetNextView() as View;
                while (view != null)
                {
                    string name = view.GetName2();
                    if (!string.IsNullOrEmpty(name))
                    {
                        viewNames.Add(name);
                        Logger.Info($"Found view: {name}");
                    }

                    // Move to next
                    view = view.GetNextView() as View;
                }

                if (viewNames.Count == 0)
                {
                    Logger.Warn("No drawing views found (after sheet).");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Exception while getting view names: {ex.Message}");
            }

            return viewNames;
        }

        public bool InsertBreakLineOnView(
            string viewName,
            double pos1,
            double pos2,
            swBreakLineOrientation_e orientation = swBreakLineOrientation_e.swBreakLineHorizontal,
            swBreakLineStyle_e style = swBreakLineStyle_e.swBreakLine_Jagged,
            int shapeIntensity = 3,
            bool breakSketchBlocks = true)
        {
            if (_drawingDoc == null)
            {
                Logger.Error("DrawingDoc is null. Cannot insert break line.");
                return false;
            }

            try
            {
                Logger.Info($"Starting break line insertion on view '{viewName}'...");

                // Activate and select the view
                bool activated = _drawingDoc.ActivateView(viewName);
                if (!activated)
                {
                    Logger.Warn($"Failed to activate view '{viewName}'.");
                    return false;
                }

                bool selected = _modelDoc.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                if (!selected)
                {
                    Logger.Warn($"Failed to select view '{viewName}'.");
                    return false;
                }

                var selectionMgr = (SelectionMgr)_modelDoc.SelectionManager;
                var swView = (View)selectionMgr.GetSelectedObject6(1, -1);

                if (swView == null)
                {
                    Logger.Warn("Failed to get selected view object.");
                    return false;
                }

                // Insert the break line using InsertBreak3
                var swBreakLineObj = swView.InsertBreak3(
                    (int)orientation,
                    pos1,
                    pos2,
                    (int)style,
                    shapeIntensity,
                    breakSketchBlocks);

                if (swBreakLineObj == null)
                {
                    Logger.Warn("InsertBreak3 returned null. Break line not created.");
                    return false;
                }

                Logger.Success($"Break line inserted successfully on view '{viewName}'.");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"Exception during break line insertion: {ex.Message}");
                return false;
            }
        }

        /*
        public bool CreateBrokenView(
              string viewName,
              double gapSize,
              swBreakLineOrientation_e orientation = swBreakLineOrientation_e.swBreakLineHorizontal,
              swBreakLineStyle_e style = swBreakLineStyle_e.swBreakLine_Jagged)
        {
            if (_drawingDoc == null)
            {
                Logger.Error("DrawingDoc is null. Cannot create a broken view.");
                return false;
            }

            try
            {
                Logger.Info($"Attempting to create broken view on '{viewName}'...");

                // Clear all selections before selecting the view
                _modelDoc.ClearSelection2(true);

                // Activate and select the target view
                if (!_drawingDoc.ActivateView(viewName))
                {
                    Logger.Warn($"Failed to activate view '{viewName}'.");
                    return false;
                }

                if (!_modelDoc.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0))
                {
                    Logger.Warn($"Failed to select view '{viewName}'.");
                    return false;
                }


                var selectionMgr = (SelectionMgr)_modelDoc.SelectionManager;
                var swView = (View)selectionMgr.GetSelectedObject6(1, -1);

                if (swView == null)
                {
                    Logger.Warn("Failed to get the selected View object.");
                    return false;
                }

                // Get bounding box outline
                double[] box = (double[])swView.GetOutline();
                if (box == null || box.Length < 4)
                {
                    Logger.Warn("Failed to get view outline for break positions.");
                    return false;
                }

                // Calculate break positions depending on orientation
                double pos1, pos2;
                if (orientation == swBreakLineOrientation_e.swBreakLineHorizontal)
                {
                    // Use Y coordinates
                    pos1 = box[1] + (box[3] - box[1]) * 0.3;
                    pos2 = box[1] + (box[3] - box[1]) * 0.7;
                }
                else
                {
                    // Use X coordinates
                    pos1 = box[0] + (box[2] - box[0]) * 0.3;
                    pos2 = box[0] + (box[2] - box[0]) * 0.7;
                }

                Logger.Info($"Break positions calculated: {pos1:F4} to {pos2:F4}");

                // Insert break line
                var swBreakLine = (BreakLine)swView.InsertBreak3(
                    (int)orientation,
                    pos1,
                    pos2,
                    (int)style,
                    2,  // Shape intensity (for jagged)
                    true // Break sketch blocks
                );

                if (swBreakLine == null)
                {
                    Logger.Warn("InsertBreak3 returned null. Break line not created.");
                    return false;
                }

                // Execute the break
                _drawingDoc.BreakView();

                // Rebuild
                _modelDoc.ClearSelection2(true);
                _modelDoc.EditRebuild3();

                Logger.Success($"Successfully created broken view on '{viewName}'.");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"An exception occurred while creating the broken view: {ex.Message}");
                return false;
            }
        }
        */

        public bool CreateBrokenView(
            string viewName,
            double breakLine1Pos,
            double breakLine2Pos,
            double gapSize,
            swBreakLineOrientation_e orientation = swBreakLineOrientation_e.swBreakLineHorizontal,
            swBreakLineStyle_e style = swBreakLineStyle_e.swBreakLine_Jagged)
        {
            if (_drawingDoc == null)
            {
                Logger.Error("DrawingDoc is null. Cannot create a broken view.");
                return false;
            }

            try
            {
                Logger.Info($"Attempting to create broken view on '{viewName}'...");

                // Step 1: Activate and select the target view
                if (!_drawingDoc.ActivateView(viewName))
                {
                    Logger.Warn($"Failed to activate view '{viewName}'.");
                    return false;
                }

                if (!_modelDoc.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0))
                {
                    Logger.Warn($"Failed to select view '{viewName}'.");
                    return false;
                }

                var selectionMgr = (SelectionMgr)_modelDoc.SelectionManager;
                var swView = (View)selectionMgr.GetSelectedObject6(1, -1);

                if (swView == null)
                {
                    Logger.Warn("Failed to get the selected View object.");
                    return false;
                }

                // Step 2: Use InsertBreak to create the PAIR of lines for a true broken view.
                var swBreakLine = (BreakLine)swView.InsertBreak((int)orientation, breakLine1Pos, breakLine2Pos, (int)style);

                if (swBreakLine == null)
                {
                    Logger.Warn("IView::InsertBreak returned null. Failed to create break line pair.");
                    return false;
                }

                // Step 3: Execute the command to actually break the view.
                _drawingDoc.BreakView();

                // Step 4: Set the desired gap size.
                //swView.SetBreakLineGap(gapSize);

                // Step 5: Rebuild to apply changes.
                _modelDoc.ClearSelection2(true);
                _modelDoc.EditRebuild3();

                Logger.Success($"Successfully created broken view on '{viewName}'.");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"An exception occurred while creating the broken view: {ex.Message}");
                return false;
            }
        }

    }
}
