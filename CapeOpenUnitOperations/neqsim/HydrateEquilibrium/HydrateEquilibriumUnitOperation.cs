using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using CAPEOPEN110;
using Microsoft.Win32;

namespace CapeOpenUnitOperations
{
    [Serializable]
    [Guid("A33373D4-C594-4544-8116-C02C032F8501")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("NeqSim.CapeOpen.Hydrate")]
    public class HydrateEquilibriumUnitOperation : UnitOperationBaseClass
    {
        private readonly double temperature = 20.0;

        public HydrateEquilibriumUnitOperation()
        {
            componentDescription = "Hydrate Equilibrium Unit";
            componentName = "Hydrate Equilibrium Unit";
            reportNames = new string[1];
        }

        public override object reports => reportNames;

        public override string selectedReport
        {
            set => reportNames[0] = value;
            get => reportNames[0];
        }

        public override object simulationContext
        {
            set => SimulationContextLocal = value;
        }

        public override object parameters => parameterCollection;


        public override object ports => portCollection;

        public override CapeValidationStatus ValStatus
        {
            get
            {
                var ValStatus2 = CapeValidationStatus.CAPE_NOT_VALIDATED;
                return ValStatus2;
                // throw new NotImplementedException();
            }
        }


        public override string ComponentDescription
        {
            get => componentDescription;
            set => componentDescription = value;
        }

        public override string ComponentName
        {
            get => componentName;
            set => componentName = value;
        }

        public override void ProduceReport(ref string report)
        {
            report = "Hydrate equilibrium temperature " + (neqsimService.getTemperature() - 273.15) + "\n";
            report += "Pressure " + neqsimService.getPressure();
        }


        public override void Initialize()
        {
            reportNames[0] = "Hydrate info";
            portCollection = new PortCollectionBaseClass();
            ((PortCollectionBaseClass) portCollection).addPort("Port1", CapePortDirection.CAPE_INLET);
            //((PortCollectionBaseClass)portCollection).addPort("Port2", CapePortDirection.CAPE_OUTLET);

            parameterCollection = new ParameterCollectionBaseClass();
            ((ParameterCollectionBaseClass) parameterCollection).AddParameter("Hydrate equilibrium temperature",
                "Hydrate temperature", CapeParamMode.CAPE_OUTPUT, new ParameterRealSpec(), 1);
            ((ParameterBaseClass) parameterCollection.Item("Hydrate equilibrium temperature")).value = temperature;
            // parameterCollection = new ParameterCollectionBaseClass();

            //   portCollection = new PortCollectionBaseClass();
            //   ((PortCollectionBaseClass)portCollection).addPort("Port1", CapePortDirection.CAPE_INLET);
            //   ((PortCollectionBaseClass)portCollection).addPort("Port2", CapePortDirection.CAPE_OUTLET);

            //var salmons = new List<PortBaseClass>();
            //PortBaseClass port1 = new PortBaseClass();
            //salmons.Add(port1);

            //PortBaseClass[] ports = new PortBaseClass[1];
            //ports[0] = port1;
            //object tempPort = ports;
            //localPorts2 = (ICapeCollection)salmons;
            // this.ProduceReport
        }

        public override void Calculate()
        {
            var inputMat = (ICapeThermoMaterial) ((PortBaseClass) portCollection.Item("Port1")).connectedObject;
            double temperature = 0;
            double pressure = 0;
            object composition = null;

            try
            {
                inputMat.GetOverallTPFraction(ref temperature, ref pressure, ref composition);
                neqsimService.setTPFraction(temperature, pressure / 1.0e5, (double[]) composition);
                neqsimService.HydrateEquilibriumTemperature();
            }
            catch (Exception e)
            {
                var w = e.Message;
                throw e;
            }

            temperature = neqsimService.getTemperature() - 273.15;
            ((ParameterBaseClass) parameterCollection.Item("Hydrate equilibrium temperature")).value = temperature;
        }

        public override void Edit()
        {
            ((ParameterBaseClass) parameterCollection.Item("Hydrate equilibrium temperature")).value = temperature;
           // double pressure = 0;
        }

        public override void Terminate()
        {
            ((ParameterBaseClass) parameterCollection.Item("Hydrate equilibrium temperature")).value = temperature;
        }


        public override bool Validate(ref string errText)
        {
            if (portCollection == null) return false;
            if ((ICapeThermoMaterial) ((PortBaseClass) portCollection.Item("Port1")).connectedObject == null)
            {
                errText = "input streem need to be connected";
                return false;
            }

            return true;
        }

        # region COM Registration

        [ComRegisterFunction]
        public new static void RegisterFunction(Type t)
        {
            const string ICapeOpenComponent0 = "{678C09A1-7D66-11D2-A67D-00105A42887F}";
            const string ICapeOpenComponent = "{678C09A5-7D66-11D2-A67D-00105A42887F}";
            const string ICapeOpenComponent1 = "{4150C28A-EE06-403f-A871-87AFEC38A249}";
            const string ICapeOpenComponent2 = "{4667023A-5A8E-4CCA-AB6D-9D78C5112FED}";


            // const String ICapeOpenThermo = "{CF51E383-0110-4ed8-ACB7-B50CFDE6908E}";
            try
            {
                MemberInfo inf = typeof(HydrateEquilibriumUnitOperation);
                var CLSID = Registry.ClassesRoot.OpenSubKey("CLSID", RegistryKeyPermissionCheck.ReadWriteSubTree,
                    RegistryRights.CreateSubKey);
                var attributes = inf.GetCustomAttributes(typeof(GuidAttribute), false);
                var guid = "{" + ((GuidAttribute) attributes[0]).Value + "}";

                var key = CLSID.OpenSubKey(guid, true);

                var CapeDescription = key.CreateSubKey("CapeDescription", RegistryKeyPermissionCheck.ReadWriteSubTree);

                CapeDescription.SetValue("About", "NeqSim Hydrate Equilibrium Unit");
                CapeDescription.SetValue("CapeVersion", "1.0");
                CapeDescription.SetValue("ComponentVersion", "1.0-0");
                CapeDescription.SetValue("Name", "NeqSim Hydrate Equilibrium Unit");
                CapeDescription.SetValue("Description", "NeqSim Hydrate Equilibrium Unit");
                CapeDescription.SetValue("HelpUrl", "http://143.97.83.56:8080/NeqSimWiki/en/NeqSim_Wiki");
                CapeDescription.SetValue("VendorUrl", "NeqSim Thermo");
                CapeDescription.SetValue("Description",
                    "NeqSim is a process simulation and design tool used in oil and gas production. NeqSim thermodynamic and unit operaions can by used in 3rd part simulation tools supporting the Cape Open interface.");
                key.SetValue("", "Tet UnitOp Obect -NET");

                var ImplementedCategories = key.OpenSubKey("Implemented Categories", true);
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
        public new  static void UnregisterFunction(Type t)
        {
            //CapeOpenRegistration.UnRegisterFunction(typeof(ThermoPackages));
        }

        # endregion
    }
}