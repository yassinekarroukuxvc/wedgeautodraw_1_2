using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Core.Interfaces;

public interface INoteService
{
    public bool InsertDimensionNote(DataStorage position, string[] wedgeKeys, string header, DrawingData drawingData, NamedDimensionValues wedgeDimensions);
    public bool InsertOverlayCalibrationNote(string calibrationValueMicrons, double squareSideInInches);
    public bool InsertCustomNoteAtPosition(string noteText, DataStorage position);
    public bool InsertCustomNoteAsTable(string noteText, DataStorage position);
}
