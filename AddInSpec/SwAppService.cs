using SolidWorks.Interop.sldworks;

namespace AddInSpec
{
    public static class SwAppService
    {
        public static SldWorks SwApp { get; private set; }

        public static void Initialize(SldWorks app)
        {
            SwApp = app;
        }
    }
}
