using System;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet28
    {
        private void Sheet28_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet28_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.calcButton.Click += new System.EventHandler(this.button1_Click);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.activteSHeet);
            this.Startup += new System.EventHandler(this.Sheet28_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet28_Shutdown);

        }

        #endregion

        private void activteSHeet()
        {
            var rangeClear = Range["A2", "D100"];
            rangeClear.Clear();
            rangeClear.Font.Color = ColorTranslator.ToOle(Color.DarkGray);

            var rangeBLack = Range["B2", "B200"];
            rangeBLack.Font.Color = ColorTranslator.ToOle(Color.Black);

            var a = NeqSimThermoSystem.getThermoSystem().getPhase(0).getNumberOfComponents() + 1;
            var range = Range["B2", "B" + a];
            var i = 0;
            string number;
            foreach (Range r in range.Cells)
                if (string.IsNullOrEmpty((string) r.Text))
                {
                    number = (i + 2).ToString();
                    Range["A" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i)
                        .getComponentName();
                    Range["B" + number].Value2 =
                        NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getz() * 100;
                    Range["C" + number].Value2 =
                        NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getMolarMass() * 1000.0;
                    Range["D" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i)
                        .getNormalLiquidDensity();
                    i++;
                }

            number = (i + 3).ToString();
            Range["A" + number].Value2 = "Total %";
            Range["B" + number].Value2 = "=sum(B2:B" + (i + 1) + ")";
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }
    }
}