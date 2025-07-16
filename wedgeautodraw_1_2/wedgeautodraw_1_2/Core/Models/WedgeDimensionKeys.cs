using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using wedgeautodraw_1_2.Core.Enums;

namespace wedgeautodraw_1_2.Core.Models;

public static class WedgeDimensionKeys
{
    public static readonly Dictionary<WedgeType, HashSet<string>> TypeToKeys = new()
    {
        [WedgeType.CKVD] = new HashSet<string>
            {
                "TL", "TD", "TDF", "BA", "ISA", "FA", "GR", "GA", "GD", "FL", "F", "FX", "B", "VW", "E", "FR", "BR",
                "X", "VR", "W", "FRX", "BRX", "FL_groove_angle"
            },
        [WedgeType.COB] = new HashSet<string>
            {
                "TL", "TD", "TDF", "BA", "ISA", "RA", "T", "ERW", "ERD", "CA", "FD", "FL", "FLER", "FRO",
                "FH", "HA", "FNA", "H", "MB", "W"
            },
        
    };

    public static readonly Dictionary<WedgeType, HashSet<string>> TypeToAngleKeys = new()
    {
        [WedgeType.CKVD] = new HashSet<string>
            {
                "ISA", "FA", "BA", "GA", "FL_groove_angle"
            },
        [WedgeType.COB] = new HashSet<string>
            {
                "ISA", "BA", "RA", "CA", "HA", "FNA"
            },
        
    };
}
