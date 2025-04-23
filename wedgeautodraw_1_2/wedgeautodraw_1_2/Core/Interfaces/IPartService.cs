using wedgeautodraw_1_2.Core.Models;

namespace wedgeautodraw_1_2.Core.Interfaces;

public interface IPartService
{
    void OpenPart(string partPath);
    void UpdateEquations(string equationFilePath);
    void SetEngravedText(string text);
    void ApplyTolerances(DynamicDataContainer dimensions);
    void ToggleSketchVisibility(string sketchName, bool visible);
    void EnableSolveOrder(bool enable);
    void EnableAutoRebuild(bool enable);
    void Rebuild();
    void Save(bool close = false);
    void Reopen(string partPath);
    void Unlock();
}
