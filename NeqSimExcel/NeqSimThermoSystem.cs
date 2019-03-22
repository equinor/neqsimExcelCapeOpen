using neqsim.thermo.system;

namespace NeqSimExcel
{
    public static class NeqSimThermoSystem
    {
        private static SystemInterface thermoSystem = new SystemSrkEos(273, 1.0);

        public static string LocalFilePath { get; set; } = null;

        // static NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter test = null;

        public static SystemInterface getThermoSystem()
        {
            return thermoSystem;
        }

        public static void setThermoSystem(SystemInterface system)
        {
            thermoSystem = system;
        }

        public static void setDatabasePath(string databaseType, string databasePath)
        {


            if (databaseType.Equals("MSAccess"))
            {
                neqsim.util.database.NeqSimDataBase.setDataBaseType("MSAccessUCanAccess");
                neqsim.util.database.NeqSimDataBase.dataBasePath = databasePath;
            }
            if (databaseType.Equals("mySQL"))
            {
                neqsim.util.database.NeqSimDataBase.setDataBaseType("mySQL");
            }
    }
    }
}