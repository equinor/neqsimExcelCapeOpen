using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet28
    {
        private void Sheet28_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet28_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
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
            Excel.Range rangeClear = this.Range["A2", "D100"];
            rangeClear.Clear();
            rangeClear.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);

            Excel.Range rangeBLack = this.Range["B2", "B200"];
            rangeBLack.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

            int a = NeqSimThermoSystem.getThermoSystem().getPhase(0).getNumberOfComponents() + 1;
            Excel.Range range = this.Range["B2", "B" + a.ToString()];
            int i = 0;
            string number;
            foreach (Excel.Range r in range.Cells)
            {
                if (String.IsNullOrEmpty((string)r.Text))
                {
                    number = (i + 2).ToString();
                    this.Range["A" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getComponentName();
                    this.Range["B" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getz() * 100;
                    this.Range["C" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getMolarMass() * 1000.0;
                    this.Range["D" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getNormalLiquidDensity();
                    i++;
                }
            }
            number = (i + 3).ToString();
            this.Range["A" + number].Value2 = "Total %";
            this.Range["B" + number].Value2 = "=sum(B2:B" + (i + 1) + ")";

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
