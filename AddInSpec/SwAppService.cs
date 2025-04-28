using SolidWorks.Interop.sldworks;


namespace AddInSpec
{
    public static class SwAppService
    {
        public static SldWorks SwApp { get; private set; }

        public static void Initialize(ISldWorks app)
        {
            SwApp = (SldWorks)app;
        }
    }
}
