﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using CAPEOPEN110;
using Microsoft.Win32;

namespace CapeOpenThermo
{
    [Serializable]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("F15C825A-3108-43B6-B503-9F58447B3920")]
    [ProgId("Statoil.CapeOpen.Local")]
    public class ThermoPackageManagerCO11local : ICapeIdentification, ICapeThermoPropertyPackageManager, ICapeUtilities,
        IDisposable
    {
        public object SimulationContextLocal;

        public string ComponentDescription
        {
            get => ComponentDescription;
            set => ComponentDescription = value;
        }

        public string ComponentName
        {
            get => ComponentName;
            set => ComponentName = value;
        }

        public object GetPropertyPackageList()
        {
            var names = new List<string>();
            var filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";

            var d = new DirectoryInfo(fullPath);
            var Files = d.GetFiles("*.neqsim");
            foreach (var file in Files) names.Add(file.Name.Replace(".neqsim", ""));

            return names.ToArray();
        }

        public object GetPropertyPackage(string package)
        {
            var filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";
            var filename = fullPath + "\\" + package + ".neqsim";
            var test = new NeqSimNETClientCO11(filename, package);
            return test;
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

        public virtual object simulationContext
        {
            set => SimulationContextLocal = value;
        }

        public virtual object parameters => throw new NotImplementedException();

        public void Dispose()
        {
        }

        # region COM Registration

        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            const string ICapeOpenComponent = "{678c09a1-7d66-11d2-a67d-00105a42887f}";
            const string ICapeOpenThermo = "{CF51E383-0110-4ed8-ACB7-B50CFDE6908E}";

            try
            {
                MemberInfo inf = typeof(ThermoPackageManagerCO11local);
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
                        "Failed to access registry for NeqSim-Cape Open (GUID). You have to have adminitration rights to do this!");
                var CapeDescription = key.CreateSubKey("CapeDescription", RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (CapeDescription == null)
                    throw new COMException(
                        "Failed to access registry for NeqSim-Cape Open (CapeDescription). You have to have adminitration rights to do this!");
                CapeDescription.SetValue("About", "NeqSim Thermo Cape Open Package");
                CapeDescription.SetValue("CapeVersion", "1.1");
                CapeDescription.SetValue("ComponentVersion", "1.0-0");
                CapeDescription.SetValue("Name", "NeqSim Thermo local");
                CapeDescription.SetValue("HelpUrl", "http://143.97.83.56:8080/NeqSimWiki/en/NeqSim_Wiki");
                CapeDescription.SetValue("VendorUrl", "NeqSim Thermo");
                CapeDescription.SetValue("Description",
                    "NeqSim is a process simulation and design tool used in oil and gas production. NeqSim thermodynamic and unit operaions can by used in 3rd part simulation tools supporting the Cape Open interface.");
                key.SetValue("", "Tet Obect -NET");

                var ImplementedCategories = key.OpenSubKey("Implemented Categories", true);
                if (ImplementedCategories == null)
                    throw new COMException(
                        "Failed to access registry for NeqSim Cape Open(Implemented Categories). You have to have adminitration rights to do this!");
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

        #endregion
    }
}