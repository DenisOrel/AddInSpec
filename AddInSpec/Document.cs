using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;

namespace AddInSpec
{
    internal class Document : SwAddin
    {
        private static SldWorks _swApp;
        private static ModelDoc2 _swModel;
        private static AssemblyDoc _swAssy;
        private static int nRetVal;

        private Document()
        {
            _swApp = SwAppService.SwApp;
        }
        private void CheckDocument()
        {
            var swModel = (ModelDoc2)_swApp.ActiveDoc;

            if (swModel == null)
            {
                // Handle case when no document is open
                _swApp.SendMsgToUser2("No document is open.", (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
                return;
            }

            if (_swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                // Handle assembly document
                _swAssy = (AssemblyDoc)_swModel;

                DistinctComponents();
                Get();
            }
            else 
            {
                // Handle case when active document is not assembly
                _swApp.SendMsgToUser2("Active document is not assembly.", (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
            }
        }

        public List<PropList> GetDistinctPartComponents()
        {
            var groupedComponents = new Dictionary<string, Component2>();

            var compList = new List<PropList>();

            nRetVal = _swAssy.ResolveAllLightWeightComponents(false);

            var vComponents = (object[])_swAssy.GetComponents(true);

            foreach (Component2 component in vComponents)
            {
                if (component.GetSuppression2() == (int)swComponentSuppressionState_e.swComponentSuppressed) continue;

                compList.Add(new PropList
                {
                    Configuration = component.ReferencedConfiguration,
                    Filename = component.GetPathName(),
                    SwModel = component.GetModelDoc2()
                });
            }

            return compList;
        }

        private List<PropList> DistinctComponents()
        {
            var list = new List<PropList>();
            GetDistinctPartComponents();

            // TODO: Create a new list to store the distinct components and add quantity

            return list;
        }

        public void Get()
        {
            //todo: Get properties from the list of components

            foreach (var comp in DistinctComponents())
            {
                GetCustomProp(comp.SwModel, "Обозначение", comp.Configuration);
            }
        }

        private string GetCustomProp(ModelDoc2 SwModel, string fieldName, string config = "")
        {
            var ext = SwModel.Extension;
            var customPropMan = ext.CustomPropertyManager[config];

            customPropMan.Get6(FieldName: fieldName, UseCached: false, ValOut: out string value, ResolvedValOut: out string resolvedValue, WasResolved: out bool unit, LinkToProperty: out bool type);

            return resolvedValue;
        }
    }
}
