using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using CAPEOPEN110;
using Microsoft.Win32;

namespace CapeOpenThermo
{
    [Serializable]
    [Guid("7C244400-C27B-4F95-ABED-38AF9920B23C")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("Equinor.CapeOpen")]
    public class ThermoPackageManagerCO11 : ICapeIdentification, ICapeThermoPropertyPackageManager
    {
        public string componentDescription = "NeqSim Thermo  Package online";
        public string componentName = "NeqSim Thermo  Package online";

        public object SimulationContextLocal;

        public virtual string ComponentDescription
        {
            get => componentDescription;
            set => componentDescription = value;
        }

        public virtual string ComponentName
        {
            get => componentName;
            set => componentName = value;
        }

        public object GetPropertyPackageList()
        {

            var names = new List<string>();

            return names.ToArray();
        }

        public object GetPropertyPackage(string package)
        {
            return new NeqSimNETClientCO11(package);
        }

        # region COM Registration

        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            const string ICapeOpenComponent = "{678c09a1-7d66-11d2-a67d-00105a42887f}";

            const string ICapeOpenThermo = "{CF51E383-0110-4ed8-ACB7-B50CFDE6908E}";

            try
            {
                MemberInfo inf = typeof(ThermoPackageManagerCO11);
                var CLSID = Registry.ClassesRoot.OpenSubKey("CLSID", RegistryKeyPermissionCheck.ReadWriteSubTree,
                    RegistryRights.CreateSubKey);
                if (CLSID == null)
                    throw new COMException(
                        "Failed to access registry for NeqSim-Cape Open(CLSID). You have to have adminitration rights to do this!");
                var attributes = inf.GetCustomAttributes(typeof(GuidAttribute), false);
                var guid = "{" + ((GuidAttribute) attributes[0]).Value + "}";
                var key = CLSID.OpenSubKey(guid, true);
                if (key == null)
                    throw new COMException(
                        "Failed to access registry for NeqSim-Cape Open (GUID). You need to have administrator rights to do this!");
                var CapeDescription = key.CreateSubKey("CapeDescription", RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (CapeDescription == null)
                    throw new COMException(
                        "Failed to access registry for NeqSim-Cape Open (CapeDescription). You have to have adminitration rights to do this!");
                CapeDescription.SetValue("About", "NeqSim Thermo Cape Open Package");
                CapeDescription.SetValue("CapeVersion", "1.1");
                CapeDescription.SetValue("ComponentVersion", "1.0-1");
                CapeDescription.SetValue("Name", "NeqSim Thermo");
                CapeDescription.SetValue("HelpUrl", "https://equinor.github.io/neqsimhome/");
                CapeDescription.SetValue("VendorUrl", "NeqSim Thermo");
                CapeDescription.SetValue("Description",
                    "NeqSim is a process simulation and design tool used in oil and gas production. NeqSim thermodynamic and unit operations can by used in 3rd part simulation tools supporting the Cape Open interface.");
                key.SetValue("", "Tet Obect -NET");

                var ImplementedCategories = key.OpenSubKey("Implemented Categories", true);
                if (ImplementedCategories == null)
                    throw new COMException(
                        "Failed to access registry for NeqSim Cape Open(Implemented Categories). You need to have administrator rights to do this!");
                ImplementedCategories.CreateSubKey(ICapeOpenThermo);
                ImplementedCategories.CreateSubKey(ICapeOpenComponent);


                Console.WriteLine("Registering NeqSim Cape Open Objects...ok");
            }
            catch (Exception e)
            {
                e.ToString();
                Console.WriteLine("Registering NeqSim Cape Open Objects...failed");
                Console.WriteLine(e.ToString());
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
        }

        # endregion
    }
}