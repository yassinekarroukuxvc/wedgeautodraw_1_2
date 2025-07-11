using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services;

public class NoteService : INoteService
{
    private readonly SldWorks _swApp;
    private ModelDoc2 _swModel;
    private DrawingDoc _swDrawing;
    private ModelDocExtension _swModelExt;
    private CustomPropertyManager _swCustProps;
    private string _drawingPath;

    private int _error = 0;
    private int _warning = 0;
    public NoteService(SldWorks swApp, ModelDoc2 swModel)
    {
        _swApp = swApp;
        _swModel = swModel;
        _swDrawing = _swModel as DrawingDoc;
    }
    public bool InsertDimensionNote(DataStorage position, string[] wedgeKeys, string header, DrawingData drawingData, NamedDimensionValues wedgeDimensions)
    {
        if (_swModel == null)
        {
            Logger.Warn("ModelDoc2 is null. Cannot insert note.");
            return false;
        }

        try
        {
            var validLines = new List<string>();

            foreach (var key in wedgeKeys)
            {
                if (wedgeDimensions.TryGet(key, out var dataStorage) && dataStorage != null)
                {
                    double valueInch = dataStorage.GetValue(Unit.Inch);
                    double upperTolInch = dataStorage.GetTolerance(Unit.Inch, "+");
                    double lowerTolInch = dataStorage.GetTolerance(Unit.Inch, "-");

                    if (!double.IsNaN(valueInch))
                    {
                        string inchStr = TrimLeadingZero(valueInch.ToString("F5"));
                        string tolStrInch = FormatTolerance(upperTolInch, lowerTolInch, 4, true);

                        string line = $"{key} = {inchStr}{tolStrInch}";
                        validLines.Add(line.Trim());
                    }
                }
            }

            if (validLines.Count == 0)
            {
                Logger.Warn("No valid dimensions to insert as a note.");
                return false;
            }

            string noteText = string.Join("\n", validLines);
            double[] pos = position.GetValues(Unit.Meter);

            object noteObj = _swModel.InsertNote(noteText);
            if (noteObj == null)
            {
                Logger.Warn("InsertNote returned null.");
                return false;
            }

            var note = noteObj as Note;
            if (note == null)
            {
                Logger.Warn("Failed to cast note object to Note.");
                return false;
            }

            Annotation annotation = (Annotation)note.GetAnnotation();
            if (annotation == null)
            {
                Logger.Warn("Failed to get annotation from note.");
                return false;
            }

            annotation.SetPosition2(pos[0], pos[1], 0.0);
            annotation.Layer = "Annotaion";

            TextFormat format = (TextFormat)note.GetTextFormat();
            format.CharHeight = 0.0040;
            format.TypeFaceName = "Arial";
            format.Bold = true;
            format.Italic = false;
            format.Underline = false;

            bool applied = annotation.SetTextFormat(0, false, format);
            if (!applied)
            {
                Logger.Warn("SetTextFormat failed.");
            }

            note.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);

            Logger.Info("Dimension note inserted and formatted successfully.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception during note insertion: {ex.Message}");
            return false;
        }
    }

    private static string TrimLeadingZero(string input)
    {
        return input.StartsWith("0.") ? input.Substring(1) : input;
    }

    private static string FormatTolerance(double upper, double lower, int precision, bool inch)
    {
        string fmt = inch ? "F" + precision : "F3";

        if ((upper == 0 && lower == 0) ||
            double.IsNaN(upper) || double.IsNaN(lower))
        {
            return " (REF)";
        }

        if (lower == 0)
            return $"+{upper.ToString(fmt)}";

        if (upper == 0)
            return $"-{lower.ToString(fmt)}";

        if (upper == lower)
            return $"±{upper.ToString(fmt)}";

        return $"+{upper.ToString(fmt)} -{lower.ToString(fmt)}";
    }


    public bool InsertOverlayCalibrationNote(string calibrationValueMicrons, double squareSideInInches)
    {
        if (_swModel == null)
        {
            Logger.Warn("ModelDoc2 is null. Cannot insert calibration note.");
            return false;
        }

        try
        {
            // Compute bottom-right position of the square
            double sideM = squareSideInInches * 0.0254;
            Sheet sheet = (Sheet)_swDrawing.GetCurrentSheet();
            double sheetWidth = 0.0, sheetHeight = 0.0;
            sheet.GetSize(ref sheetWidth, ref sheetHeight);

            double centerX = sheetWidth / 2.0;
            double centerY = sheetHeight / 2.0;

            double posX = centerX + sideM / 2.0 - 0.007;
            double posY = centerY - sideM / 2.0 + 0.005;

            string noteText = $"{calibrationValueMicrons}µm";

            object noteObj = _swModel.InsertNote(noteText);
            if (noteObj is not Note note)
            {
                Logger.Warn("Failed to insert overlay calibration note.");
                return false;
            }

            Annotation annotation = (Annotation)note.GetAnnotation();
            if (annotation == null)
            {
                Logger.Warn("Failed to get annotation from note.");
                return false;
            }

            annotation.SetPosition2(posX, posY, 0.0);

            TextFormat format = (TextFormat)note.GetTextFormat();
            format.CharHeight = 0.0025;
            format.TypeFaceName = "Arial";
            format.Bold = false;
            format.Italic = false;
            format.Underline = false;

            annotation.SetTextFormat(0, false, format);
            note.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);

            Logger.Info("Overlay Calibration note inserted successfully.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception inserting calibration note: {ex.Message}");
            return false;
        }
    }
    public bool InsertCustomNoteAtPosition(string noteText, DataStorage position)
    {
        if (_swModel == null)
        {
            Logger.Warn("ModelDoc2 is null. Cannot insert custom note.");
            return false;
        }

        try
        {
            double[] pos = position.GetValues(Unit.Meter);

            object noteObj = _swModel.InsertNote(noteText);
            if (noteObj is not Note note)
            {
                Logger.Warn("InsertNote returned null or failed to cast.");
                return false;
            }

            Annotation annotation = (Annotation)note.GetAnnotation();
            if (annotation == null)
            {
                Logger.Warn("Failed to get annotation from note.");
                return false;
            }

            annotation.SetPosition2(pos[0], pos[1], 0.0);
            annotation.Layer = "Annotaion";
            note.SetBalloon(
            (int)swBalloonStyle_e.swBS_Square,  // Style: rectangle
            8                                       // Size: 2 = medium
        );
            

            TextFormat format = (TextFormat)note.GetTextFormat(); // <-- Explicit cast
            format.CharHeight = 0.0035;
            format.TypeFaceName = "Arial";
            format.Bold = false;
            format.Italic = false;
            format.Underline = false;
            

            bool applied = annotation.SetTextFormat(0, false, format);
            if (!applied)
            {
                Logger.Warn("SetTextFormat failed.");
            }

            note.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);

            Logger.Info($"Custom note inserted successfully: \"{noteText}\"");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception during custom note insertion: {ex.Message}");
            return false;
        }
    }
    public bool InsertCustomNoteAsTable(string noteText, DataStorage position)
    {
        if (_swModel == null || _swDrawing == null)
        {
            Logger.Warn("ModelDoc2 or DrawingDoc is null. Cannot insert table.");
            return false;
        }

        try
        {
            double[] pos = position.GetValues(Unit.Meter);
            int rows = 1;

            // Insert a 1-row, 1-column generic table
            TableAnnotation table = _swDrawing.InsertTableAnnotation2(
                false,
                pos[0],
                pos[1],
                1,
                "",
                rows,
                1);

            if (table == null)
            {
                Logger.Warn("Failed to insert table.");
                return false;
            }

            // Set text content
            table.Text[0, 0] = noteText;

            // Set column width (e.g., 80 mm)
            table.SetColumnWidth(0, 0.10, (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);

            // Set row height (e.g., 20 mm)
            table.SetRowHeight(0, 0.02, (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);

            // Remove grid lines
            table.GridLineWeight = (int)swLineWeights_e.swLW_NONE;

            // Apply layer and position
            Annotation annotation = (Annotation)table.GetAnnotation();
            if (annotation != null)
            {
                annotation.Layer = "Annotaion";
                annotation.SetPosition2(pos[0], pos[1], 0.0);
            }

            // Text formatting
            TextFormat format = (TextFormat)table.GetTextFormat();
            format.CharHeight = 0.005;
            format.TypeFaceName = "Arial";
            format.Bold = false;
            format.Italic = false;
            format.Underline = false;

            table.SetTextFormat(false, format);
            Logger.Info($"Custom text table inserted as note: \"{noteText}\"");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Exception inserting custom table: {ex.Message}");
            return false;
        }
    }


}
