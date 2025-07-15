using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using wedgeautodraw_1_2.Core.Enums;

namespace wedgeautodraw_1_2.Core.Models;

public static class WedgeDimensionKeys
{
    public static readonly Dictionary<WedgeType, List<string>> Keys = new()
    {
        { WedgeType.CKVD, new List<string> { "TL", "W", "FL", "TD", "TDF", "F", "W", "FR", "BR", "GD", "GR", "B", "E", "X", "VR", "VW", "FX" } },
        { WedgeType.COB, new List<string> { "TLA", "Width", "Angle" } },
        { WedgeType.OSG7, new List<string> { "Length", "Diameter", "Chamfer" } },
        
    };
}
