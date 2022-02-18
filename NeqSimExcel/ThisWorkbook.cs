using System;
using neqsim.thermo.system;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class ThisWorkbook
    {
        private SystemInterface thermoSystem = new SystemSrkEos(298, 10);

        private void ThisWorkbook_Startup(object sender, EventArgs e)
        {
            Application.UseSystemSeparators = true;
        }

        private void ThisWorkbook_Shutdown(object sender, EventArgs e)
        {
        }


        public SystemInterface getThermoSystem()
        {
            return thermoSystem;
        }

        public void setThermoSystem(SystemInterface system)
        {
            thermoSystem = system;
        }


        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisWorkbook_Startup;
            Shutdown += ThisWorkbook_Shutdown;
        }

        #endregion
    }
}