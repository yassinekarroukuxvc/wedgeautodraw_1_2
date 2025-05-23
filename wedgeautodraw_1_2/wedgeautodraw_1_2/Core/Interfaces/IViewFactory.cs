namespace wedgeautodraw_1_2.Core.Interfaces;

public interface IViewFactory
{
    IViewService CreateView(string viewName);
}
