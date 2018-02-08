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
    public partial class Sheet9
    {
        private void Sheet9_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet9_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button2.Click += new System.EventHandler(this.button2_Click);
            this.buttonTCPComega.Click += new System.EventHandler(this.buttonTCPComega_Click);
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.activteSHeet);
            this.Startup += new System.EventHandler(this.Sheet9_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet9_Shutdown);

        }

        #endregion

        public void button1_Click_1(object sender, EventArgs e)
        {
            Excel.Range rangeClear = this.Range["A2", "H100"];
            rangeClear.Clear();
            rangeClear.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);

            Excel.Range rangeBLack = this.Range["B2", "B200"];
            rangeBLack.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

            int a = NeqSimThermoSystem.getThermoSystem().getPhase(0).getNumberOfComponents()+1;
            Excel.Range range = this.Range["B2", "B" + a.ToString()];
            int i=0;
            string number;
            foreach (Excel.Range r in range.Cells)
            {
                if (String.IsNullOrEmpty((string)r.Text))
                {
                    number = (i + 2).ToString();
                    this.Range["A" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getComponentName();
                    this.Range["B" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getz()*100;
                    this.Range["C" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getMolarMass()*1000.0;
                    this.Range["D" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getNormalLiquidDensity();
                    this.Range["E" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getTC()-273.15;
                    this.Range["F" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getPC();
                    this.Range["G" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getAcentricFactor();
                    i++;
                }
            }
            number = (i + 3).ToString();
            this.Range["A" + number].Value2 = "Total %";
            this.Range["B" + number].Value2 = "=sum(B2:B" + (i + 1) + ")";

        }

        public void button2_Click(object sender, EventArgs e)
        {
            this.Range["J7"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            this.Range["J7"].Value2 = "Updating fluid composition....please wait...";
            int a = NeqSimThermoSystem.getThermoSystem().getPhase(0).getNumberOfComponents() + 1;
            Excel.Range range = this.Range["B2", "B" + a.ToString()];
            int i = 0;
            double[] fractions = new double[a-1];
            
            foreach (Excel.Range r in range.Cells)
            {
                if (!String.IsNullOrEmpty((string)r.Text))
                {
                    fractions[i++] = r.Value2;
                }
            }

            Excel.Range namerange = this.Range["A2", "A" + a.ToString()];
            string[] names = new string[a - 1];
             i = 0;

            foreach (Excel.Range r in namerange.Cells)
            {
                if (!String.IsNullOrEmpty((string)r.Text))
                {
                    for (int k = 0; k < NeqSimThermoSystem.getThermoSystem().getMaxNumberOfPhases(); k++)
                    {
                        NeqSimThermoSystem.getThermoSystem().getPhase(k).getComponent(i).setComponentName(r.Value2);
                    }
                        names[i++] = r.Value2;
                }
            }
            NeqSimThermoSystem.getThermoSystem().setComponentNames(names);

            NeqSimThermoSystem.getThermoSystem().setMolarComposition(fractions);
            NeqSimThermoSystem.getThermoSystem().init(0);
            NeqSimThermoSystem.getThermoSystem().init(1);
            this.Range["J7"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
            button1_Click_1(sender, e);
            this.Range["J7"].Value2 = "done...fluid composition updated";
        }

        private void buttonTCPComega_Click(object sender, EventArgs e)
        {
            this.Range["J7"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            this.Range["J7"].Value2 = "Updating TC, PC, omega....please wait...";
            int a = NeqSimThermoSystem.getThermoSystem().getPhase(0).getNumberOfComponents() + 1;
            Excel.Range range = this.Range["A2", "A" + a.ToString()];
            int i = 1;
           
            foreach (Excel.Range r in range.Cells)
            {
                NeqSimThermoSystem.getThermoSystem().init(0);
                i++;
                for (int k = 0; k < NeqSimThermoSystem.getThermoSystem().getMaxNumberOfPhases(); k++)
                {
                    NeqSimThermoSystem.getThermoSystem().getPhase(k).getComponent(((string)r.Text)).setTC(this.Range[("E" + i.ToString())].Value2 + 273.15);
                    NeqSimThermoSystem.getThermoSystem().getPhase(k).getComponent(((string)r.Text)).setPC(this.Range[("F" + i.ToString())].Value2);
                    NeqSimThermoSystem.getThermoSystem().getPhase(k).getComponent(((string)r.Text)).setAcentricFactor(this.Range[("G" + i.ToString())].Value2);
                }
            }
            
            NeqSimThermoSystem.getThermoSystem().init(0);
            NeqSimThermoSystem.getThermoSystem().init(1);
            this.Range["J7"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
            button1_Click_1(sender, e);
            this.Range["J7"].Value2 = "done...updated critical parameters";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ((thermo.phase.PhaseEosInterface)NeqSimThermoSystem.getThermoSystem().getPhase(0)).displayInteractionCoefficients("");
        }

        private void activteSHeet()
        {
            Excel.Range rangeClear = this.Range["A2", "H100"];
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
                    this.Range["E" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getTC() - 273.15;
                    this.Range["F" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getPC();
                    this.Range["G" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getAcentricFactor();
                    i++;
                }
            }
            number = (i + 3).ToString();
            this.Range["A" + number].Value2 = "Total %";
            this.Range["B" + number].Value2 = "=sum(B2:B" + (i + 1) + ")";

        }
    }
}
