using System;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet26
    {
        private void Sheet26_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet26_Shutdown(object sender, EventArgs e)
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
            Startup += Sheet26_Startup;
            Shutdown += Sheet26_Shutdown;
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var thermoSystem = NeqSimThermoSystem.getThermoSystem(); //.clone();
            thermoSystem.init(0);
            thermoSystem.init(1);

            var Cnumb = 5;
            for (var i = 2; i <= 100; i++)
            {
                string a = Range["A" + i].Value2;
                if (a != null)
                    thermoSystem.addTBPfraction(Range["A" + i].Value2, 1.0e-10, Range["C" + i].Value2 / 1000.0,
                        Range["D" + i].Value2 / 1000.0);
            }

            thermoSystem.autoSelectMixingRule();
            thermoSystem.init(0);
            thermoSystem.init(1);
            thermoSystem.setMultiPhaseCheck(true);
            thermoSystem.resetPhysicalProperties();
            thermoSystem.initPhysicalProperties();
            NeqSimThermoSystem.setThermoSystem(thermoSystem);
            Globals.Sheet9.button1_Click_1(sender, e);
        }
    }
}