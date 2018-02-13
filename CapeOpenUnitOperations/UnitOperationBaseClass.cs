using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using CAPEOPEN110;
using Microsoft.Win32;
using System.Reflection;
using System.Data;
using System.Collections;
using System.Security.Principal;
using NeqSimNET;


namespace CapeOpenUnitOperations
{
    [Serializable]
    [Guid("CD1CB03C-F23B-4472-9440-FBB1AD6E9F2C")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("NeqSim.CapeOpen.BaseOp")]

    public class UnitOperationBaseClass : CAPEOPEN110.ICapeUnit, ICapeIdentification, ICapeUtilities, ICapeUnitReport
    {
        public string componentDescription = "NeqSim Unit Operation";
        public string componentName = "NeqSim Unit Sep";
        public string[] reportNames = new string[1];
       
        public Object SimulationContextLocal = null;
        public ICapeCollection portCollection = null;
        public ICapeCollection parameterCollection;

        public int fluidNumber = 2341;

        public NeqSimNETService neqsimService = null;

        public UnitOperationBaseClass(){
            neqsimService = new NeqSimNETService();
            fluidNumber = neqsimService.getPackageID();
            neqsimService.readFluidFromGQIT(fluidNumber);
        }
       
        public virtual void ProduceReport(ref String name)
        {
        }

        public virtual Object reports
        {
            get
            {
                return reportNames;
            }
          
        }

        public virtual String selectedReport
        {
            set
            {
                reportNames[0] = value;
            }
            get
            {
                return reportNames[0];
            }

        }

        public virtual void Initialize()
        {
            reportNames[0] = "report11";

            portCollection = new PortCollectionBaseClass();
            ((PortCollectionBaseClass)portCollection).addPort("Port1", CapePortDirection.CAPE_INLET);
            ((PortCollectionBaseClass)portCollection).addPort("Port2", CapePortDirection.CAPE_OUTLET);

            parameterCollection = new ParameterCollectionBaseClass();
            ((ParameterCollectionBaseClass)parameterCollection).AddParameter("NeqSimUnitID", "Online unit ID from NeqSim", CapeParamMode.CAPE_INPUT, new ParameterIntSpec(), 1);
            ((ParameterCollectionBaseClass)parameterCollection).AddParameter("NeqSimUnitID2", "Online unit ID from NeqSim", CapeParamMode.CAPE_INPUT, new ParameterIntSpec(), 1);
           
        }

        public virtual void Calculate()
        {
            ICapeThermoMaterial inputMat = (ICapeThermoMaterial)((PortBaseClass)portCollection.Item("Port1")).connectedObject;
            double temperature = 0;
            double pressure = 0;
            object composition = null;
            try
            {
               inputMat.GetOverallTPFraction(ref temperature, ref pressure, ref composition);
            }
            catch (Exception e)
            {
                String w = e.Message;
                throw e;
            }

            
                ICapeThermoMaterial outputMat = (ICapeThermoMaterial)((PortBaseClass)portCollection.Item("Port2")).connectedObject;
                outputMat.CopyFromMaterial(inputMat);
                outputMat.SetOverallProp("pressure", "", 30e5);
        }

        public virtual void Edit()
        {
            double pressure = 0;
        }

        public virtual void Terminate()
        {
            double pressure = 0;
        }

        public virtual Object simulationContext
        {
            set
            {
                SimulationContextLocal = value;
                //CapeInterface = value;
                //  throw new NotImplementedException();
            }
        }

        public virtual Object parameters
        {
            get
            {
    //            return null;
                return parameterCollection;
            }
        }


        public virtual Object ports
        {
            get
            {
                return portCollection;
            }
        }


        public virtual bool Validate(ref String errText)
        {
            if (portCollection == null) return false;
            if ((ICapeThermoMaterial)((PortBaseClass)portCollection.Item("Port1")).connectedObject == null || (ICapeThermoMaterial)((PortBaseClass)portCollection.Item("Port2")).connectedObject == null)
            {
                errText = "input streem need to be connected";
                return false;
            }
            else return true;
        }

        public virtual CAPEOPEN110.CapeValidationStatus ValStatus
        {
            get
            {
                CAPEOPEN110.CapeValidationStatus ValStatus2 = CAPEOPEN110.CapeValidationStatus.CAPE_NOT_VALIDATED;
                return ValStatus2;  
                // throw new NotImplementedException();
            }
        }




        public virtual String ComponentDescription
        {
            get
            {
                return componentDescription;
                // throw new NotImplementedException();
            }
            set
            {
                componentDescription = value;
                //  throw new NotImplementedException();
            }
        }

        public virtual String ComponentName
        {
            get
            {
                return componentName;
                // throw new NotImplementedException();
            }
            set
            {
                componentName = value;
                //throw new NotImplementedException();
            }
        }

        # region COM Registration
        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
         
            const string ICapeOpenComponent0 = "{678C09A1-7D66-11D2-A67D-00105A42887F}";
            const string ICapeOpenComponent = "{678C09A5-7D66-11D2-A67D-00105A42887F}";
            const string ICapeOpenComponent1 = "{4150C28A-EE06-403f-A871-87AFEC38A249}";
            const string ICapeOpenComponent2 = "{4667023A-5A8E-4CCA-AB6D-9D78C5112FED}";
            
                
           // const String ICapeOpenThermo = "{CF51E383-0110-4ed8-ACB7-B50CFDE6908E}";
            try
            {
                MemberInfo inf = typeof(UnitOperationBaseClass);
                RegistryKey CLSID = Registry.ClassesRoot.OpenSubKey("CLSID", RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.CreateSubKey);
                object[] attributes = inf.GetCustomAttributes(typeof(GuidAttribute), false);
                string guid = "{" + ((GuidAttribute)attributes[0]).Value + "}";

                RegistryKey key = CLSID.OpenSubKey(guid, true);

                RegistryKey CapeDescription = key.CreateSubKey("CapeDescription", RegistryKeyPermissionCheck.ReadWriteSubTree);

                CapeDescription.SetValue("About", "NeqSim Unit Operation Cape Open Package");
                CapeDescription.SetValue("CapeVersion", "1.0");
                CapeDescription.SetValue("ComponentVersion", "1.0-0");
                CapeDescription.SetValue("Name", "NeqSim Unit Op");
                CapeDescription.SetValue("Description", "NeqSim Unit Op");
                CapeDescription.SetValue("HelpUrl", "http://143.97.83.56:8080/NeqSimWiki/en/NeqSim_Wiki");
                CapeDescription.SetValue("VendorUrl", "NeqSim Thermo");
                CapeDescription.SetValue("Description", "NeqSim is a process simulation and design tool used in oil and gas production. NeqSim thermodynamic and unit operaions can by used in 3rd part simulation tools supporting the Cape Open interface.");
                key.SetValue("", "Tet UnitOp Obect -NET");

                RegistryKey ImplementedCategories = key.OpenSubKey("Implemented Categories", true);
                ImplementedCategories.CreateSubKey(ICapeOpenComponent0);
                ImplementedCategories.CreateSubKey(ICapeOpenComponent);
                ImplementedCategories.CreateSubKey(ICapeOpenComponent1);
                ImplementedCategories.CreateSubKey(ICapeOpenComponent2);
            }
            catch (Exception e)
            {
                e.ToString();
            }

            //CapeOpenRegistration.Register(typeof(ThermoPackages));
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            //CapeOpenRegistration.UnRegisterFunction(typeof(ThermoPackages));
        }
        # endregion
    }

    
}
