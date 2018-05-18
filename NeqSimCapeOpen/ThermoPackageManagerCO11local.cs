using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using CAPEOPEN110;
using Microsoft.Win32;
using System.Reflection;
using System.IO;

namespace CapeOpenThermo
{
    [Serializable]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("F15C825A-3108-43B6-B503-9F58447B3920")]
    [ProgId("Statoil.CapeOpen.Local")]

    public class ThermoPackageManagerCO11local : ICapeIdentification, ICapeThermoPropertyPackageManager, ICapeUtilities, IDisposable
    {
        public Object SimulationContextLocal = null;

        public object GetPropertyPackageList()
        {
            List<string> names = new List<string>();
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";

            DirectoryInfo d = new DirectoryInfo(@fullPath);
            FileInfo[] Files = d.GetFiles("*.neqsim");
            foreach (FileInfo file in Files)
            {
                names.Add(file.Name.Replace(".neqsim", ""));
            }

            return names.ToArray();
        }

        public object GetPropertyPackage(String package)
        {
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";
            string filename = fullPath + "\\" + package + ".neqsim";
            NeqSimNETClientCO11 test = new NeqSimNETClientCO11(filename, package);
            return test;
        }

        # region COM Registration
        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            const string ICapeOpenComponent = "{678c09a1-7d66-11d2-a67d-00105a42887f}";
            const String ICapeOpenThermo = "{CF51E383-0110-4ed8-ACB7-B50CFDE6908E}";

            try
            {
                MemberInfo inf = typeof(ThermoPackageManagerCO11local);
                RegistryKey CLSID = Registry.ClassesRoot.OpenSubKey("CLSID", RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.CreateSubKey);
                if (CLSID == null) throw new COMException("Failed to access registry for NeqSim-Cape Open(CLSID). You have to have adminitration rights to do this!");
                object[] attributes = inf.GetCustomAttributes(typeof(GuidAttribute), false);
                string guid = "{" + ((GuidAttribute)attributes[0]).Value + "}";
                RegistryKey key = CLSID.OpenSubKey(guid, true);
                if (key == null) throw new COMException("Failed to access registry for NeqSim-Cape Open (GUID). You have to have adminitration rights to do this!");
                RegistryKey CapeDescription = key.CreateSubKey("CapeDescription", RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (CapeDescription == null) throw new COMException("Failed to access registry for NeqSim-Cape Open (CapeDescription). You have to have adminitration rights to do this!");
                CapeDescription.SetValue("About", "NeqSim Thermo Cape Open Package");
                CapeDescription.SetValue("CapeVersion", "1.1");
                CapeDescription.SetValue("ComponentVersion", "1.0-0");
                CapeDescription.SetValue("Name", "NeqSim Thermo local");
                CapeDescription.SetValue("HelpUrl", "http://143.97.83.56:8080/NeqSimWiki/en/NeqSim_Wiki");
                CapeDescription.SetValue("VendorUrl", "NeqSim Thermo");
                CapeDescription.SetValue("Description", "NeqSim is a process simulation and design tool used in oil and gas production. NeqSim thermodynamic and unit operaions can by used in 3rd part simulation tools supporting the Cape Open interface.");
                key.SetValue("", "Tet Obect -NET");

                RegistryKey ImplementedCategories = key.OpenSubKey("Implemented Categories", true);
                if (ImplementedCategories == null) throw new COMException("Failed to access registry for NeqSim Cape Open(Implemented Categories). You have to have adminitration rights to do this!");
                ImplementedCategories.CreateSubKey(ICapeOpenThermo);
                ImplementedCategories.CreateSubKey(ICapeOpenComponent);


                System.Console.WriteLine("Registering NeqSim Cape Open Objects...ok");
            }
            catch (Exception e)
            {
                e.ToString();
                System.Console.WriteLine("Registering NeqSim Cape Open Objects...failed");
                Console.WriteLine(e.ToString());
            }

        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
        }
        #endregion

        public String ComponentDescription
        {
            get
            {
                return ComponentDescription;
            }
            set
            {
                ComponentDescription = value;
            }
        }

        public string ComponentName
        {
            get
            {
                return ComponentName;
            }
            set
            {
                ComponentName = value;
            }
        }

        public void Initialize()
        {

        }

        public void Terminate()
        {

        }

        public void Edit()
        {

        }

        public virtual Object simulationContext
        {
            set
            {
                SimulationContextLocal = value;
            }
        }

        public virtual Object parameters
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public void Dispose()
        {
          
        }
    }
}
