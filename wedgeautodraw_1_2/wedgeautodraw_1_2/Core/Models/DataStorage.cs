using wedgeautodraw_1_2.Core.Enums;

namespace wedgeautodraw_1_2.Core.Models;

public class DataStorage
{
    private double val = double.NaN;
    private double upperTol = double.NaN;
    private double lowerTol = double.NaN;

    private string valStr = "NaN";
    private string upperTolStr = "NaN";
    private string lowerTolStr = "NaN";

    private double[] valArr = Array.Empty<double>();
    private Unit? dimensionUnit;

    public DataStorage(double value)
    {
        val = value;
        valStr = value.ToString();
    }

    public DataStorage(string value, string upper, string lower)
    {
        valStr = value;
        upperTolStr = upper;
        lowerTolStr = lower;

        val = double.TryParse(value, out double parsedVal) ? parsedVal : double.NaN;
        upperTol = double.TryParse(upper, out double parsedUpper) ? parsedUpper : double.NaN;
        lowerTol = double.TryParse(lower, out double parsedLower) ? parsedLower : double.NaN;
    }

    public DataStorage(double[] values)
    {
        valArr = values ?? Array.Empty<double>();
    }

    public double[] GetValues(Unit unit)
    {
        return unit switch
        {
            Unit.Millimeter => valArr,
            Unit.Meter => ConvertAll(valArr, v => v / 1000),
            Unit.Inch => ConvertAll(valArr, v => v / 25.4),
            Unit.Radian => ConvertAll(valArr, v => v * Math.PI / 180),
            Unit.Degree => valArr,
            _ => Array.Empty<double>(),
        };
    }

    public double GetValue(Unit unit)
    {
        return unit switch
        {
            Unit.Millimeter => val,
            Unit.Meter => val / 1000,
            Unit.Inch => val / 25.4,
            Unit.Radian => val * Math.PI / 180,
            Unit.Degree => val,
            _ => double.NaN,
        };
    }

    public double GetTolerance(Unit unit, string sign)
    {
        return sign switch
        {
            "+" => unit switch
            {
                Unit.Millimeter => upperTol,
                Unit.Meter => upperTol / 1000,
                Unit.Inch => upperTol / 25.4,
                Unit.Radian => upperTol * Math.PI / 180,
                Unit.Degree => upperTol,
                _ => double.NaN,
            },
            "-" => unit switch
            {
                Unit.Millimeter => lowerTol,
                Unit.Meter => lowerTol / 1000,
                Unit.Inch => lowerTol / 25.4,
                Unit.Radian => lowerTol * Math.PI / 180,
                Unit.Degree => lowerTol,
                _ => double.NaN,
            },
            _ => double.NaN,
        };
    }

    private double[] ConvertAll(double[] input, Func<double, double> converter)
    {
        double[] result = new double[input.Length];
        for (int i = 0; i < input.Length; i++)
            result[i] = converter(input[i]);
        return result;
    }

    public void SetUnit(Unit unit) => dimensionUnit = unit;
}
