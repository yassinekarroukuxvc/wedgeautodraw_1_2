using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

                    double valueMm = dataStorage.GetValue(Unit.Millimeter);
                    double upperTolMm = dataStorage.GetTolerance(Unit.Millimeter, "+");
                    double lowerTolMm = dataStorage.GetTolerance(Unit.Millimeter, "-");

                    if (!double.IsNaN(valueInch))
                    {
                        bool isRef = IsRef(upperTolInch, lowerTolInch);

                        string tolStrInch = (!double.IsNaN(upperTolInch) && !double.IsNaN(lowerTolInch) && !isRef)
                            ? FormatTolerance(upperTolInch, lowerTolInch, 4, true)
                            : "";

                        string tolStrMm = (!double.IsNaN(upperTolMm) && !double.IsNaN(lowerTolMm) && !isRef)
                            ? FormatTolerance(upperTolMm, lowerTolMm, 4, false)
                            : "";

                        string inchStr = !double.IsNaN(valueInch) ? TrimLeadingZero(valueInch.ToString("F4")) : "";
                        string mmStr = !double.IsNaN(valueMm) ? valueMm.ToString("F3") : "";

                        string line = $"{key} = {inchStr} {tolStrInch}".Trim();

                        if (!string.IsNullOrWhiteSpace(mmStr))
                            line += $" [{mmStr} {tolStrMm}]";

                        if (isRef)
                            line += " (REF)";

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

            // Set position
            Annotation annotation = (Annotation)note.GetAnnotation();
            if (annotation == null)
            {
                Logger.Warn("Failed to get annotation from note.");
                return false;
            }

            annotation.SetPosition2(pos[0], pos[1], 0.0);

            // Customize text format
            TextFormat format = (TextFormat)note.GetTextFormat();
            format.CharHeight = 0.003; // 0.5 mm
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

    private static bool IsRef(double upper, double lower)
    {
        return upper == 0 && lower == 0;
    }

    private static string FormatTolerance(double upper, double lower, int precision, bool inch)
    {
        string fmt = inch ? "F" + precision : "F3";
        return $"±{upper.ToString(fmt)}";
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

            double posX = centerX + sideM / 2.0 - 0.007; // shift slightly inside
            double posY = centerY - sideM / 2.0 + 0.005; // shift slightly up

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
            format.CharHeight = 0.0025; // 0.5 mm
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
}
