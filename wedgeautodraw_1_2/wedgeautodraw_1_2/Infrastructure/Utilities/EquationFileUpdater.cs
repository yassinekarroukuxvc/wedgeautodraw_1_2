using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using wedgeautodraw_1_2.Core.Enums;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Utilities
{
    public static class EquationFileUpdater
    {
        public static void UpdateEquationFile(string equationFilePath, WedgeData wedge)
        {
            if (!File.Exists(equationFilePath))
            {
                Logger.Warn($"Equation file not found: {equationFilePath}");
                return;
            }

            var originalLines = File.ReadAllLines(equationFilePath).ToList();
            var outputLines = new List<string>();
            var dimensionKeys = wedge.Dimensions.GetAll().Keys.ToHashSet();
            var updatedKeys = new HashSet<string>();

            var dimensionRegex = new Regex("^\"(?<key>[^\"]+)\"\\s*=.*$", RegexOptions.Compiled);

            foreach (var line in originalLines)
            {
                var match = dimensionRegex.Match(line);
                if (match.Success)
                {
                    string key = match.Groups["key"].Value;
                    if (dimensionKeys.Contains(key))
                    {
                        var data = wedge.Dimensions[key];
                        string valueStr = data.GetValue(Unit.Millimeter).ToString("0.#####", CultureInfo.InvariantCulture);
                        string unitStr = IsAngle(key) ? "deg" : "mm";

                        outputLines.Add($"\"{key}\"={valueStr}{unitStr}");
                        updatedKeys.Add(key);
                        continue;
                    }
                }

                outputLines.Add(line); // Keep original line
            }

            // Append any missing dimensions from the wedge
            foreach (var kvp in wedge.Dimensions.GetAll())
            {
                string key = kvp.Key;
                if (updatedKeys.Contains(key)) continue;

                var data = kvp.Value;
                string valueStr = data.GetValue(Unit.Millimeter).ToString("0.#####", CultureInfo.InvariantCulture);
                string unitStr = IsAngle(key) ? "deg" : "mm";

                outputLines.Add($"\"{key}\"={valueStr}{unitStr}");
            }

            try
            {
                File.WriteAllLines(equationFilePath, outputLines);
                Logger.Success($"Equation file updated at: {equationFilePath}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to write equation file: {ex.Message}");
            }
        }

        private static bool IsAngle(string key)
        {
            return key is "ISA" or "FA" or "BA" or "GA" or "FL_groove_angle";
        }
    }
}
