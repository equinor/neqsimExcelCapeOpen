using System;
using Microsoft.Office.Interop.Excel;
using thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet2
    {
        private void Sheet2_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet2_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            button1.Click += button1_Click;
            Startup += Sheet2_Startup;
            Shutdown += Sheet2_Shutdown;
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            thermoSystem.setTemperature(273.75);

            var range = Range["A2", "A100"];
            var number = 1;
            foreach (Range r in range.Cells)
            {
                var text = (string) r.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    number++;
                    thermoSystem.setPressure(r.Value2);
                    var ops = new ThermodynamicOperations(thermoSystem);
                    ops.dewPointTemperatureFlash();
                    var textVar = "B" + number;
                    Range[textVar].Value2 = thermoSystem.getTemperature() - 273.15;
                }
            }
        }
    }
}