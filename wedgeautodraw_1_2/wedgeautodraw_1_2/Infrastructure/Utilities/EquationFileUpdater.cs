using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Linq;
using System.Collections.Generic;

public static class EquationFileUpdater
{
    public static void UpdateEquationFile(string equationFilePath, WedgeData wedge)
    {
        if (!File.Exists(equationFilePath))
        {
            Logger.Warn($"Equation file not found: {equationFilePath}");
            return;
        }

        var encoding = GetFileEncoding(equationFilePath);
        var originalLines = File.ReadAllLines(equationFilePath, encoding).ToList();
        var outputLines = new List<string>();

        var dimensionKeys = wedge.Dimensions.GetAll().Keys.ToHashSet();

        var wedgeType = wedge.WedgeType;
        var angleKeys = WedgeDimensionKeys.TypeToAngleKeys.ContainsKey(wedgeType)
            ? WedgeDimensionKeys.TypeToAngleKeys[wedgeType]
            : new HashSet<string>();

        var dimensionRegex = new Regex("^\"(?<key>[^\"]+)\"\\s*=.*$", RegexOptions.Compiled);

        bool overlayCalibrationUpdated = false;

        foreach (var line in originalLines)
        {
            var match = dimensionRegex.Match(line);
            if (match.Success)
            {
                string key = match.Groups["key"].Value;

                if (key == "EngravingStart" && wedge.Dimensions.ContainsKey("TL"))
                {
                    outputLines.Add("\"EngravingStart\" = \"TL\" * 0.45");
                    continue;
                }

                if (key == "overlay_calibration1")
                {
                    string overlayStr = wedge.OverlayScaling.ToString("0.#####", CultureInfo.InvariantCulture);
                    outputLines.Add($"\"overlay_calibration1\" = {overlayStr}");
                    overlayCalibrationUpdated = true;
                    continue;
                }

                if (dimensionKeys.Contains(key))
                {
                    var data = wedge.Dimensions[key];
                    string valueStr = data.GetValue(Unit.Millimeter).ToString("0.#####", CultureInfo.InvariantCulture);
                    string unitStr = angleKeys.Contains(key) ? "deg" : "mm";

                    outputLines.Add($"\"{key}\"={valueStr}{unitStr}");
                    continue;
                }
            }

            // Keep original line if no update
            outputLines.Add(line);
        }

        if (!overlayCalibrationUpdated)
        {
            string overlayStr = wedge.OverlayScaling.ToString("0.#####", CultureInfo.InvariantCulture);
            outputLines.Add($"\"overlay_calibration1\" = {overlayStr}");
        }

        try
        {
            File.WriteAllText(equationFilePath, string.Join("\r\n", outputLines), encoding);
            Logger.Success($"Equation file updated at: {equationFilePath}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to write equation file: {ex.Message}");
        }
    }

    private static Encoding GetFileEncoding(string filePath)
    {
        using var reader = new StreamReader(filePath, true);
        if (reader.Peek() >= 0)
        {
            reader.Read(); // Trigger encoding detection
        }
        return reader.CurrentEncoding;
    }

    public static void EnsureAllEquationsExist(ModelDoc2 model, WedgeData wedge)
    {
        var mgr = (EquationMgr)model.GetEquationMgr();
        int count = mgr.GetCount();

        var existingKeys = new HashSet<string>();
        for (int i = 0; i < count; i++)
        {
            string eq = mgr.Equation[i];
            string key = eq.Split('=')[0].Trim().Trim('"');
            existingKeys.Add(key);
        }

        model.ClearSelection2(true);
        model.EditRebuild3();

        var wedgeType = wedge.WedgeType;
        var angleKeys = WedgeDimensionKeys.TypeToAngleKeys.ContainsKey(wedgeType)
            ? WedgeDimensionKeys.TypeToAngleKeys[wedgeType]
            : new HashSet<string>();

        foreach (var kvp in wedge.Dimensions.GetAll())
        {
            string key = kvp.Key;
            if (existingKeys.Contains(key)) continue;

            double value = kvp.Value.GetValue(Unit.Millimeter);

            string equation = angleKeys.Contains(key)
                ? $"\"{key}\" = {value.ToString("0.#####", CultureInfo.InvariantCulture)}deg"
                : $"\"{key}\" = {value.ToString("0.#####", CultureInfo.InvariantCulture)}";

            int idx = mgr.Add3(
                -1,
                equation,
                true,
                (int)swInConfigurationOpts_e.swThisConfiguration,
                null
            );

            if (idx >= 0)
                Logger.Info($"Added equation to part: {equation}");
            else
                Logger.Warn($"[EquationMgr] Add3 failed for: {equation} (key: {key})");
        }

        model.EditRebuild3();
    }
}
