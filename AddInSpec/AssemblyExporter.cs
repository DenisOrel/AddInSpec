using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;

namespace AddInSpec
{
    internal class AssemblyExporter
    {
        private readonly SldWorks _swApp = SwAppService.SwApp;

        public void ExportAssembly(string[] attributes, string outputPath)
        {
            try
            {
                ModelDoc2 model = _swApp.ActiveDoc;

                if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                    throw new InvalidDataException("Active document is not an assembly.");

                var configurationName = model.ConfigurationManager.ActiveConfiguration.Name;
                var fileName = Path.GetFileName(model.GetPathName());

                var assemblyInfo = new Assembly
                {
                    Configuration = configurationName,
                    Filename = fileName,
                    Properties = GetCustomProperties(model, attributes),
                    Components = TraverseTopLevelComponents((AssemblyDoc)model, attributes)
                };

                var root = new Root { Assembly = assemblyInfo };

                var json = JsonConvert.SerializeObject(root, Formatting.Indented);

                if (!Directory.Exists(Path.GetDirectoryName(outputPath)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
                }

                File.WriteAllText(outputPath, json);
            }
            catch (Exception e)
            {
                _swApp.SendMsgToUser(e.StackTrace);
                throw;
            }
        }

        private List<Component> TraverseTopLevelComponents(AssemblyDoc assembly, string[] attributes)
        {
            var componentGroups = new Dictionary<string, (Component component, int count)>();

            var comps = (object[])assembly.GetComponents(true); // Top level only
            if (comps == null) return new List<Component>();

            foreach (Component2 comp in comps)
            {
                if (comp.GetSuppression2() == (int)swComponentSuppressionState_e.swComponentLightweight)
                {
                    //_swApp.SendMsgToUser(comp.Name + " - swComponentLightweight");
                    var rest = comp.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved);
                    //_swApp.SendMsgToUser2(
                    //    $"Component: {comp.Name2} → SetSuppression2 result: {rest}",
                    //    (int)swMessageBoxIcon_e.swMbInformation,
                    //    (int)swMessageBoxBtn_e.swMbOk
                    //);
                }

                if (comp.IsSuppressed()) continue;
                var config = comp.ReferencedConfiguration ?? "";
                var path = comp.GetPathName() ?? "";
                var key = $"{config}|{path}";

                if (!componentGroups.ContainsKey(key))
                {
                    componentGroups[key] = (new Component
                    {
                        Configuration = config,
                        Filename = Path.GetFileName(path),
                        Quantity = "1",
                        Properties = GetCustomProperties(comp, attributes)
                    }, 1);
                }
                else
                {
                    componentGroups[key] = (componentGroups[key].component, componentGroups[key].count + 1);
                }
            }

            // Перетворення в список
            var result = new List<Component>();
            foreach (var kvp in componentGroups)
            {
                var component = kvp.Value.component;
                component.Quantity = kvp.Value.count.ToString();
                result.Add(component);
            }

            return result;
        }

        private Dictionary<string, string> GetCustomProperties(ModelDoc2 model, string[] attributes)
        {
            var props = new Dictionary<string, string>();
            if (model == null || attributes == null) return props;

            foreach (var attr in attributes)
            {
                var custPropMgr =
                    model.Extension.CustomPropertyManager[model.ConfigurationManager.ActiveConfiguration.Name];
                custPropMgr.Get6(attr, false, out _, out var resolvedValOut, out _, out _);
                if (string.IsNullOrWhiteSpace(resolvedValOut))
                {
                    custPropMgr = model.Extension.CustomPropertyManager[""];
                    custPropMgr.Get6(attr, false, out _, out resolvedValOut, out _, out _);
                }

                props[attr] = string.IsNullOrWhiteSpace(resolvedValOut) ? "empty" : resolvedValOut;
            }

            return props;
        }

        private Dictionary<string, string> GetCustomProperties(Component2 comp, string[] attributes)
        {
            var props = new Dictionary<string, string>();

            if (comp == null || attributes == null) return props;

            var model = (ModelDoc2)comp.GetModelDoc2();
            var ex = model.Extension;

            foreach (var attr in attributes)
            {
                var custPropMgr = ex.CustomPropertyManager[comp.ReferencedConfiguration];

                custPropMgr.Get6(attr, false, out _, out var resolvedValOut, out _, out _);

                if (string.IsNullOrWhiteSpace(resolvedValOut))
                {
                    custPropMgr = ex.CustomPropertyManager[""];
                    custPropMgr.Get6(attr, false, out _, out resolvedValOut, out _, out _);
                }

                props[attr] = string.IsNullOrWhiteSpace(resolvedValOut) ? "empty" : resolvedValOut;
            }

            return props;
        }
    }
}