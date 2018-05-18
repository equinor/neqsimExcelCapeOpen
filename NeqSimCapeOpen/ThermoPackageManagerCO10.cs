using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using CAPEOPEN100;
//using CAPEOPEN110;
using Microsoft.Win32;
using System.Reflection;
using System.Security.Principal;

namespace CapeOpenThermo
{
    [Serializable]
    [Guid("A68E5F20-F6F9-40D6-975B-157507C77E7C")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("Statoil.CapeOpen10")]

    public class ThermoPackageManagerCO10 : ICapeIdentification, ICapeThermoSystem
    {

        public object GetPropertyPackages()
        {

            DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter test = new DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter();
           
           string userName = WindowsIdentity.GetCurrent().Name;
            userName = userName.Replace("STATOIL-NET\\", "");
            userName = userName.Replace("WIN-NTNU-NO\\", "");
            userName = userName.ToLower();
            DatabaseConnection.NeqSimDatabaseSet.fluidinfoDataTable tt = test.GetDataBy(userName);
       
            String nametext = WindowsIdentity.GetCurrent().Name;


            List<string> names = new List<string>();
           foreach (DatabaseConnection.NeqSimDatabaseSet.fluidinfoRow row in tt.Rows)
            {
                names.Add(row.ID.ToString());
            }



            // String[] packages = { "CPApackage", "NeqSim CPA-EoS", "GERG-EoS", names.ToArray()};
            return names.ToArray();
        }

        public object ResolvePropertyPackage(String package)
        {
           return new NeqSimNETClientCO10(package); ;
        }

 

        # region COM Registration
        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            const string ICapeOpenComponent = "{678c09a1-7d66-11D2-a67d-00105a42887f}";
           const string ICapeOpenComponent3 = "{678c09a3-7d66-11D2-a67d-00105a42887f}";
            
            try
            {
                MemberInfo inf = typeof(ThermoPackageManagerCO10);
                RegistryKey CLSID = Registry.ClassesRoot.OpenSubKey("CLSID", RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.CreateSubKey);
                if (CLSID == null) throw new COMException("Failed to access registry for NeqSim-Cape Open(CLSID). You have to have adminitration rights to do this!");
                object[] attributes = inf.GetCustomAttributes(typeof(GuidAttribute), false);
                string guid = "{" + ((GuidAttribute)attributes[0]).Value + "}";
                RegistryKey key = CLSID.OpenSubKey(guid, true);
                if (key == null) throw new COMException("Failed to access registry for NeqSim-Cape Open (GUID). You have to have adminitration rights to do this!");
                RegistryKey CapeDescription = key.CreateSubKey("CapeDescription", RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (CapeDescription == null) throw new COMException("Failed to access registry for NeqSim-Cape Open (CapeDescription). You have to have adminitration rights to do this!");
                CapeDescription.SetValue("About", "NeqSim Thermo Cape Open Package");
                CapeDescription.SetValue("CapeVersion", "1.0");
                CapeDescription.SetValue("ComponentVersion", "1.0-0");
                CapeDescription.SetValue("Name", "NeqSim Thermo 10");
                CapeDescription.SetValue("HelpUrl", "http://143.97.83.56:8080/NeqSimWiki/en/NeqSim_Wiki");
                CapeDescription.SetValue("VendorUrl", "NeqSim Thermo");
                CapeDescription.SetValue("Description", "NeqSim is a process simulation and design tool used in oil and gas production. NeqSim thermodynamic and unit operaions can by used in 3rd part simulation tools supporting the Cape Open interface.");
                key.SetValue("", "Tet Obect -NET");

                RegistryKey ImplementedCategories = key.OpenSubKey("Implemented Categories", true);
                if (ImplementedCategories == null) throw new COMException("Failed to access registry for NeqSim Cape Open(Implemented Categories). You have to have adminitration rights to do this!");
                ImplementedCategories.CreateSubKey(ICapeOpenComponent3);
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
        # endregion

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
    }
}

