using System;
using Microsoft.Office.Interop.Excel;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet29
    {
        private void Sheet29_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet29_Shutdown(object sender, EventArgs e)
        {
        }


        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            this.Startup += new System.EventHandler(this.Sheet29_Startup);

        }

        #endregion

        private void button1_Click_1(object sender, EventArgs e)
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

                    // Aq. dew point:
                    ops.waterDewPointTemperatureFlash(); // How to be sure that this is aq and not hc?
                    var textVar = "B" + number;
                    Range[textVar].Value2 = thermoSystem.getTemperature() - 273.15;

                    // Ice precipitation temp:
                    thermoSystem.init(0);
                    thermoSystem.setHydrateCheck(true);
                    thermoSystem.setMultiPhaseCheck(true);
                    thermoSystem.setTemperature(278.15); // why?
                    ops.hydrateFormationTemperature(0); // structure 0 = ice
                    var textVar2 = "C" + number;
                    Range[textVar2].Value2 = thermoSystem.getTemperature() - 273.15;


                    // Hydrate formation temp:
                    thermoSystem.setSolidPhaseCheck(false);
                    thermoSystem.setHydrateCheck(true);
                    ops.hydrateFormationTemperature();
                    var textVar3 = "D" + number;
                    Range[textVar3].Value2 = thermoSystem.getTemperature() - 273.15;
                }
            }
        }
    }
}