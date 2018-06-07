using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using thermo.system;
using thermodynamicOperations;

namespace NeqSimExcel
{
    public partial class Sheet29
    {
        private void Sheet29_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet29_Shutdown(object sender, System.EventArgs e)
        {
        }

   
        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            this.Startup += new System.EventHandler(this.Sheet29_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet29_Shutdown);

        }

        #endregion

        private void button1_Click_1(object sender, EventArgs e)
        {
            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            thermoSystem.setTemperature(273.75);

            Excel.Range range = this.Range["A2", "A100"];
            int number = 1;
            foreach (Excel.Range r in range.Cells)
            {
                string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    number++;
                    thermoSystem.setPressure(r.Value2);
                    ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);

                    // Aq. dew point:
                    ops.waterDewPointTemperatureFlash(); // How to be sure that this is aq and not hc?
                    string textVar = "B" + number.ToString();
                    this.Range[textVar].Value2 = thermoSystem.getTemperature() - 273.15;

                    // Ice precipitation temp:
                    thermoSystem.init(0);
                    thermoSystem.setHydrateCheck(true);
                    thermoSystem.setMultiPhaseCheck(true);
                    thermoSystem.setTemperature(278.15); // why?
                    ops.hydrateFormationTemperature(0); // structure 0 = ice
                    string textVar2 = "C" + number.ToString();
                    this.Range[textVar2].Value2 = thermoSystem.getTemperature() - 273.15;


                    // Hydrate formation temp:
                    thermoSystem.setSolidPhaseCheck(false);
                    thermoSystem.setHydrateCheck(true);
                    ops.hydrateFormationTemperature();
                    string textVar3 = "D" + number.ToString();
                    this.Range[textVar3].Value2 = thermoSystem.getTemperature() - 273.15;


                }
            }

        }
    }
}
