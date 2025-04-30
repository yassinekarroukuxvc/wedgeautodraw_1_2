using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public static class SectionViewAdjuster
{
    public static double[] ApplyOffset(string dimName, double[] originalPos, double secv, double FL, double GD)
    {
        double offsetX = 0.0;
        double offsetY = 0.0;

        switch (dimName)
        {
            case "F":
                offsetX = 0.0;
                offsetY = -0.005;
                break;

            case "FL":
                offsetX = -0.003;
                offsetY = 0.0;
                break;

            case "FR":
                offsetX = (secv * FL / 8);
                offsetY = secv * GD / 3;
                break;

            case "BR":
                offsetX = -(secv * FL / 8);
                offsetY = secv * GD / 3;
                break;

            case "W":
                offsetX = 0.0;
                offsetY = -8;
                break;

            case "GD":
                offsetX = -2;
                offsetY = 0.0;
                break;

            default:
                offsetX = 0.0;
                offsetY = 0.0;
                break;
        }

        return new[] { originalPos[0] * 1000 + offsetX, originalPos[1] * 1000 + offsetY };
    }
}
