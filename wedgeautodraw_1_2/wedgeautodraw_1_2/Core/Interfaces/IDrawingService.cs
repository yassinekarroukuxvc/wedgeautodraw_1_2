using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Core.Interfaces;

public interface IDrawingService
{
    void OpenDrawing(string filePath);
    void SaveDrawing();
    void SaveAndCloseDrawing();
    void SaveAsPdf(string outputPath);
    void SetSummaryInformation(DrawingData drawingData);
    void SetCustomProperties(DrawingData drawingData);
    void Rebuild();
    void ZoomToFit();
    ModelDoc2 GetModel();
    void ReplaceReferencedModel(string drawingPath, string oldModelPath, string newModelPath);
}
