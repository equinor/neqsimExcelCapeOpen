using System;
using CapeOpenThermo;

namespace NeqSimRegistration
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //ThermoPackages tempVar = new ThermoPackages();
            //Object test = tempVar.GetPropertyPackageList();
            Type t = null;
            ThermoPackageManagerCO11.RegisterFunction(t);
            ThermoPackageManagerCO10.RegisterFunction(t);
            ThermoPackageManagerCO11local.RegisterFunction(t);
            Console.WriteLine("finished registering NeqSim CapeOpen2....");
            CapeOpenUnitOperations.HydrateEquilibriumUnitOperation.RegisterFunction(t);
            //CapeOpenUnitOperations.WaterDewPointUnitOperation.RegisterFunction(t);
            //  CapeOpenUnitOperations.FreezingUnitOperation.RegisterFunction(t);
            //  CapeOpenUnitOperations.ProCapMixerUnitOperation.RegisterFunction(t);
            //  CapeOpenUnitOperations.UnitOperationBaseClass.RegisterFunction(t);

            Console.WriteLine("finished registering NeqSim CapeOpen....");
           // Console.ReadLine();
            // testing..
            //testing2
            //testing3
            //testing4
        }
    }
}