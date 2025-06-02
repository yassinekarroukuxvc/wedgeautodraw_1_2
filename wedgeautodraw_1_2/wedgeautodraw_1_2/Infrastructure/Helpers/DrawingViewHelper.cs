using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public static class DrawingViewHelper
{
    public static List<string> GetAllViewNames(ModelDoc2 drawingModel)
    {
        var viewNames = new List<string>();

        if (drawingModel is not DrawingDoc drawingDoc)
        {
            Console.WriteLine("The provided model is not a drawing.");
            return viewNames;
        }

        View currentView = (View)drawingDoc.GetFirstView();

        // Skip the first view because it's the "Sheet" view
        if (currentView != null)
            currentView = (View)currentView.GetNextView();

        while (currentView != null)
        {
            string viewName = currentView.Name;
            viewNames.Add(viewName);
            currentView = (View)currentView.GetNextView();
        }

        return viewNames;
    }
}
