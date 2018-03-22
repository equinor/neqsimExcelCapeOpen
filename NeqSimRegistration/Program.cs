using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NeqSimRegistration
{
    class Program
    {
        static void Main(string[] args)
        {
            //ThermoPackages tempVar = new ThermoPackages();
            //Object test = tempVar.GetPropertyPackageList();
            Type t = null;
            CapeOpenThermo.ThermoPackageManagerCO11.RegisterFunction(t);
            CapeOpenThermo.ThermoPackageManagerCO10.RegisterFunction(t);
            System.Console.WriteLine("finished registering NeqSim CapeOpen2....");
            //CapeOpenUnitOperations.HydrateEquilibriumUnitOperation.RegisterFunction(t);
            //   CapeOpenUnitOperations.WaterDewPointUnitOperation.RegisterFunction(t);
            //  CapeOpenUnitOperations.FreezingUnitOperation.RegisterFunction(t);
            //  CapeOpenUnitOperations.ProCapMixerUnitOperation.RegisterFunction(t);
            //  CapeOpenUnitOperations.UnitOperationBaseClass.RegisterFunction(t);

            System.Console.WriteLine("finished registering NeqSim CapeOpen....");
            System.Console.ReadLine();
            // testing..
            //testing2
            //testing3
            //testing4
        }
    }
}
