using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using wedgeautodraw_1_2.Infrastructure.Helpers;

namespace wedgeautodraw_1_2.Infrastructure.Services.ViewServices;

public class ViewScaler
{
    private readonly View _swView;

    public ViewScaler(View swView)
    {
        _swView = swView;
    }

    public bool SetScale(double scale)
    {
        if (_swView == null)
        {
            Logger.Warn("Cannot set scale. View is null.");
            return false;
        }

        try
        {
            _swView.ScaleDecimal = scale;
            Logger.Info($"Set view scale to {scale:F3}.");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to set view scale: {ex.Message}");
            return false;
        }
    }

    public double GetScale()
    {
        return _swView?.ScaleDecimal ?? 1.0;
    }
}
