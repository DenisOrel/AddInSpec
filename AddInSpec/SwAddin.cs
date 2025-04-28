using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;

namespace AddInSpec
{
    [ComVisible(true)]
    [Guid("4A222941-CF4C-4A29-AB2F-C1B89D611249")]
    [DisplayName("OpenCAD")]
    [Description("Automate your work in SOLIDWORKS!")]
    public class SwAddin : SolidWorks.Interop.swpublished.SwAddin
    {
        //public SldWorks SwApp { get; private set; }
        private ISldWorks _iSwApp;

        private int _addinId;

        #region Registration

        private const string AddinKeyTemplate = @"SOFTWARE\SolidWorks\Addins\{{{0}}}";
        private const string AddinStartupKeyTemplate = @"Software\SolidWorks\AddInsStartup\{{{0}}}";
        private const string AddInTitleRegKeyName = "Title";
        private const string AddInDescriptionRegKeyName = "Description";

        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            try
            {
                string addInTitle;
                string addInDesc;

                var dispNameAtt = t.GetCustomAttributes(false).OfType<DisplayNameAttribute>().FirstOrDefault();

                if (dispNameAtt != null)
                {
                    addInTitle = dispNameAtt.DisplayName;
                }
                else
                {
                    addInTitle = t.ToString();
                }

                var descAtt = t.GetCustomAttributes(false).OfType<DescriptionAttribute>().FirstOrDefault();

                if (descAtt != null)
                {
                    addInDesc = descAtt.Description;
                }
                else
                {
                    addInDesc = t.ToString();
                }

                var addInkey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(
                    string.Format(AddinKeyTemplate, t.GUID));

                if (addInkey != null)
                {
                    addInkey.SetValue(null, 0);

                    addInkey.SetValue(AddInTitleRegKeyName, addInTitle);
                    addInkey.SetValue(AddInDescriptionRegKeyName, addInDesc);
                }

                var addInStartupkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                    string.Format(AddinStartupKeyTemplate, t.GUID));

                if (addInStartupkey != null)
                    addInStartupkey.SetValue(null, Convert.ToInt32(true),
                        Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while registering the addin: " + ex.Message);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.Registry.LocalMachine.DeleteSubKey(
                    string.Format(AddinKeyTemplate, t.GUID));

                Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(
                    string.Format(AddinStartupKeyTemplate, t.GUID));
            }
            catch (Exception e)
            {
                Console.WriteLine("Error while unregistering the addin: " + e.Message);
            }
        }

        #endregion

        public bool ConnectToSW(object thisSw, int cookie)
        {
            _iSwApp = (ISldWorks)thisSw;

            _addinId = cookie;

            //Set up callbacks
            var status = _iSwApp.SetAddinCallbackInfo2(0, this, cookie);

            SwAppService.Initialize(_iSwApp);
            
            var c = new string[1];
            _iSwApp.AddMenu((int)swDocumentTypes_e.swDocASSEMBLY, "Специфікація", 0);
            _iSwApp.AddMenuItem5((int)swDocumentTypes_e.swDocASSEMBLY, cookie, "Вивантажити дані@Специфікація", 0, "MyMenuCallback", "MyMenuEnableMethod", "My menu item", c);

            return true;
        }

        public bool DisconnectFromSW()
        {
            Marshal.ReleaseComObject(_iSwApp);
            _iSwApp = null;

            // The add-in must call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }

        public void MyMenuCallback()
        {
            var exporter = new AssemblyExporter();

            var attributes = new[]
            {
                "Обозначение", "Наименование", "Раздел",
                "Код документа", "Контора", "Масса",
                "Разработал", "Проверил"
            };

            exporter.ExportAssembly(attributes, @"C:\Temp\assembly.json");
            _iSwApp.SendMsgToUser("Complete");
        }
    }
}