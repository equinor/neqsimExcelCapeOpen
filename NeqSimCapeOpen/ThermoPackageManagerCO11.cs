using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using CAPEOPEN110;
using Microsoft.Win32;
using System.Reflection;
using System.Security.Principal;
using System.Configuration;
using System.IO;

namespace CapeOpenThermo
{
    [Serializable]
    [Guid("43F3740E-A338-49A3-AF69-B05AF1190874")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("Statoil.CapeOpen")]

    public class ThermoPackageManagerCO11 : ICapeIdentification, ICapeThermoPropertyPackageManager
    {
        public object GetPropertyPackageList2()
        {

            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(
                ConfigurationUserLevel.None);

            List<string> names = new List<string>();
            names.Add("test");
            

            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";

            DirectoryInfo d = new DirectoryInfo(@fullPath);
            FileInfo[] Files = d.GetFiles("*.neqsim");
            string str = "";
            foreach (FileInfo file in Files)
            {
                str = str + ", " + file.Name;
                names.Add(file.Name.Replace(".neqsim", ""));
            }

            return names.ToArray();
        }
        public object GetPropertyPackageList()
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
                string tempString = "";
                try
                {
                    tempString = row.TEXT.ToString();
                }
                catch(Exception e)
                {
                    tempString = "";
                    e.ToString();
                }
                finally
                {
                   

                }
                names.Add(row.ID.ToString() + " " + tempString);
               
            }
            test.Dispose();
<<<<<<< HEAD
            
=======

  

>>>>>>> f53dd0924a6d663addd8765c1b314b40e0401501
            return names.ToArray();
        }

        public object GetPropertyPackage(String package)
        {
            return new NeqSimNETClientCO11(package);
        }

        # region COM Registration
        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            const string ICapeOpenComponent = "{678c09a1-7d66-11d2-a67d-00105a42887f}";
                                               
            const String ICapeOpenThermo = "{CF51E383-0110-4ed8-ACB7-B50CFDE6908E}";

            try
            {
                MemberInfo inf = typeof(ThermoPackageManagerCO11);
               RegistryKey CLSID = Registry.ClassesRoot.OpenSubKey("CLSID", RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.CreateSubKey);
               if (CLSID == null) throw new COMException("Failed to access registry for NeqSim-Cape Open(CLSID). You have to have adminitration rights to do this!");
                object[] attributes = inf.GetCustomAttributes(typeof(GuidAttribute), false);
                string guid = "{"+ ((GuidAttribute) attributes[0]).Value + "}";
                RegistryKey key = CLSID.OpenSubKey(guid, true);
                if (key == null) throw new COMException("Failed to access registry for NeqSim-Cape Open (GUID). You have to have adminitration rights to do this!");
                RegistryKey CapeDescription = key.CreateSubKey("CapeDescription", RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (CapeDescription == null) throw new COMException("Failed to access registry for NeqSim-Cape Open (CapeDescription). You have to have adminitration rights to do this!");
               CapeDescription.SetValue("About", "NeqSim Thermo Cape Open Package");
                CapeDescription.SetValue("CapeVersion", "1.1");
                CapeDescription.SetValue("ComponentVersion", "1.0-0");
                CapeDescription.SetValue("Name", "NeqSim Thermo");
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
            catch(Exception e){
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

                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public string ComponentName
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }
    }
}
