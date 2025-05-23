
using SolidWorks.Interop.sldworks;
using wedgeautodraw_1_2.Core.Interfaces;
using wedgeautodraw_1_2.Infrastructure.Services;

namespace wedgeautodraw_1_2.Infrastructure.Factories;

public class StandardViewFactory : IViewFactory
{
    private ModelDoc2 _model;

    public StandardViewFactory(ModelDoc2 model)
    {
        _model = model;
    }

    public IViewService CreateView(string viewName)
    {
        return new ViewService(viewName, ref _model);
    }
}