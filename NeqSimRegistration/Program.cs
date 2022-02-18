using System;
using System.IO;
using CapeOpenThermo;

namespace NeqSimRegistration
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                Type t = null;
            Console.WriteLine("start registering NeqSim CapeOpen....");
            ThermoPackageManagerCO11.RegisterFunction(t);
            ThermoPackageManagerCO10.RegisterFunction(t);
            ThermoPackageManagerCO11local.RegisterFunction(t);
            CapeOpenUnitOperations.HydrateEquilibriumUnitOperation.RegisterFunction(t);
                //CapeOpenUnitOperations.WaterDewPointUnitOperation.RegisterFunction(t);
                //  CapeOpenUnitOperations.FreezingUnitOperation.RegisterFunction(t);
                //  CapeOpenUnitOperations.ProCapMixerUnitOperation.RegisterFunction(t);
                //  CapeOpenUnitOperations.UnitOperationBaseClass.RegisterFunction(t);

                Console.WriteLine("finished registering NeqSim CapeOpen....");

            }
            catch (Exception e)
            {
                Console.WriteLine("registering NeqSim CapeOpen failed: {0}", e.ToString());
                Console.ReadLine();
            }

            try
            {
                Console.WriteLine("start creating directory for neqsim fluid files....");
                string filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                filePath = filePath + "/AppData/Roaming/neqsim/fluids/";
                if (!Directory.Exists(filePath))
                {
                    DirectoryInfo di = Directory.CreateDirectory(filePath);
                }
                Console.WriteLine("finished creating directory for neqsim fluid files...." + filePath);
            }
            catch (Exception e)
            {
                Console.WriteLine("Creating directory failed: {0}", e.ToString());
                Console.ReadLine();
            }
            finally { }
        }
    }
}