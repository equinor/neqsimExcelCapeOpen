using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using thermo.system;
using thermodynamicOperations;

namespace NeqSimExcel
{
    public static class NeqSimThermoSystem
    {
        static SystemInterface thermoSystem = (SystemInterface) new SystemSrkEos(273, 1.0);
        static string localFilePath = null;

        public static string LocalFilePath { get => localFilePath; set => localFilePath = value; }

        // static NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter test = null;

        public static SystemInterface getThermoSystem()
        {
            return thermoSystem;
        }

        public static void setThermoSystem(SystemInterface system)
        {
            thermoSystem = system;
        }
    }
}
