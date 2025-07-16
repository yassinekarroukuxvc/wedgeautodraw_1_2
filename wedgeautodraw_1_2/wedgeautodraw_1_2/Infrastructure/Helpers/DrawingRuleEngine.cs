using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using wedgeautodraw_1_2.Core.Models;
using wedgeautodraw_1_2.Core.Enums;

namespace wedgeautodraw_1_2.Infrastructure.Helpers
{
    public class DrawingRuleEngine
    {
        private class RuleSet
        {
            [JsonPropertyName("conditions")]
            public List<RuleCondition> Conditions { get; set; } = new();
        }

        private class RuleCondition
        {
            [JsonPropertyName("name")]
            public string Name { get; set; }

            [JsonPropertyName("description")]
            public string Description { get; set; }

            [JsonPropertyName("drawing_type")]
            public string DrawingType { get; set; }

            [JsonPropertyName("wedge_type")]
            public string WedgeType { get; set; }

            [JsonPropertyName("if")]
            public ConditionBlock If { get; set; }

            [JsonPropertyName("update")]
            public UpdateBlock Update { get; set; }
        }

        private class ConditionBlock
        {
            [JsonPropertyName("dimension")]
            public string Dimension { get; set; }

            [JsonPropertyName("operator")]
            public string Operator { get; set; }

            [JsonPropertyName("value")]
            public double Value { get; set; }

            [JsonPropertyName("unit")]
            public string Unit { get; set; }
        }

        private class UpdateBlock
        {
            [JsonPropertyName("view_scales")]
            public Dictionary<string, double> ViewScales { get; set; }

            [JsonPropertyName("view_positions")]
            public Dictionary<string, OffsetUpdate> ViewPositions { get; set; }
        }

        private class OffsetUpdate
        {
            [JsonPropertyName("dx")]
            public double Dx { get; set; }

            [JsonPropertyName("dy")]
            public double Dy { get; set; }
        }

        private readonly RuleSet _rules;

        public DrawingRuleEngine(string jsonRuleFilePath)
        {
            if (!File.Exists(jsonRuleFilePath))
                throw new FileNotFoundException("Rules file not found", jsonRuleFilePath);

            var json = File.ReadAllText(jsonRuleFilePath);
            _rules = JsonSerializer.Deserialize<RuleSet>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
        }

        public void Apply(DrawingData drawingData, WedgeData wedgeData, DrawingType drawingType)
        {
            foreach (var rule in _rules.Conditions)
            {
                if (!string.Equals(rule.DrawingType, drawingType.ToString(), StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!string.Equals(rule.WedgeType, wedgeData.WedgeType.ToString(), StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!wedgeData.Dimensions.TryGet(rule.If.Dimension, out var dim))
                    continue;

                double actual = dim.GetValue(ParseUnit(rule.If.Unit));
                double expected = rule.If.Value;

                bool conditionMet = rule.If.Operator switch
                {
                    "<" => actual < expected,
                    ">" => actual > expected,
                    "<=" => actual <= expected,
                    ">=" => actual >= expected,
                    "==" => Math.Abs(actual - expected) < 0.0001,
                    _ => false
                };

                if (!conditionMet) continue;

                if (rule.Update.ViewScales != null)
                {
                    foreach (var kvp in rule.Update.ViewScales)
                    {
                        drawingData.ViewScales[kvp.Key] = new DataStorage(kvp.Value);
                    }

                    var frontKeys = new[] { "Front_view", "Side_view", "Top_view" };
                    bool hasFST = rule.Update.ViewScales.Keys.Any(k => frontKeys.Contains(k));

                    if (hasFST && drawingData.ViewScales.TryGet("Front_view", out var scaleData))
                    {
                        double scaleVal = scaleData.GetValue(Unit.Millimeter);
                        drawingData.TitleBlockInfo["SCALING_FRONT_SIDE_TOP_VIEW"] = scaleVal.ToString("0.###");
                    }
                }

                if (rule.Update.ViewPositions != null)
                {
                    foreach (var kvp in rule.Update.ViewPositions)
                    {
                        var current = drawingData.ViewPositions.ContainsKey(kvp.Key)
                            ? drawingData.ViewPositions[kvp.Key].GetValues(Unit.Millimeter)
                            : new[] { 0.0, 0.0 };

                        drawingData.ViewPositions[kvp.Key] = new DataStorage(new[]
                        {
                            current[0] + kvp.Value.Dx,
                            current[1] + kvp.Value.Dy
                        });
                    }
                }
            }
        }

        private static Unit ParseUnit(string unit)
        {
            return unit.ToLowerInvariant() switch
            {
                "mm" or "millimeter" => Unit.Millimeter,
                "m" or "meter" => Unit.Meter,
                "deg" or "degree" => Unit.Degree,
                _ => Unit.Millimeter
            };
        }
    }
}
