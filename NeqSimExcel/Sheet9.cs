using System;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using neqsim.thermo.phase;
using Office = Microsoft.Office.Core;


namespace NeqSimExcel
{
    public partial class Sheet9
    {
        private void Sheet9_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet9_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
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
            var rangeClear = Range["A2", "H100"];
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
                    Range["E" + number].Value2 =
                        NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getTC() - 273.15;
                    Range["F" + number].Value2 =
                        NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getPC();
                    Range["G" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i)
                        .getAcentricFactor();
                    i++;
                }

            number = (i + 3).ToString();
            Range["A" + number].Value2 = "Total %";
            Range["B" + number].Value2 = "=sum(B2:B" + (i + 1) + ")";
        }

        public void button2_Click(object sender, EventArgs e)
        {
            Range["J7"].Font.Color = ColorTranslator.ToOle(Color.Blue);
            Range["J7"].Value2 = "Updating fluid composition....please wait...";
            var a = NeqSimThermoSystem.getThermoSystem().getPhase(0).getNumberOfComponents() + 1;
            var range = Range["B2", "B" + a];
            var i = 0;
            var fractions = new double[a - 1];

            foreach (Range r in range.Cells)
                if (!string.IsNullOrEmpty((string) r.Text))
                    fractions[i++] = r.Value2;

            var namerange = Range["A2", "A" + a];
            var names = new string[a - 1];
            i = 0;

            foreach (Range r in namerange.Cells)
                if (!string.IsNullOrEmpty((string) r.Text))
                {
                    for (var k = 0; k < NeqSimThermoSystem.getThermoSystem().getMaxNumberOfPhases(); k++)
                        NeqSimThermoSystem.getThermoSystem().getPhase(k).getComponent(i).setComponentName(r.Value2);
                    names[i++] = r.Value2;
                }

            NeqSimThermoSystem.getThermoSystem().setComponentNames(names);

            NeqSimThermoSystem.getThermoSystem().setMolarComposition(fractions);
            NeqSimThermoSystem.getThermoSystem().init(0);
            NeqSimThermoSystem.getThermoSystem().init(1);
            Range["J7"].Font.Color = ColorTranslator.ToOle(Color.Green);
            button1_Click_1(sender, e);
            Range["J7"].Value2 = "done...fluid composition updated";
        }

        private void buttonTCPComega_Click(object sender, EventArgs e)
        {
            Range["J7"].Font.Color = ColorTranslator.ToOle(Color.Blue);
            Range["J7"].Value2 = "Updating TC, PC, omega....please wait...";
            var a = NeqSimThermoSystem.getThermoSystem().getPhase(0).getNumberOfComponents() + 1;
            var range = Range["A2", "A" + a];
            var i = 1;

            foreach (Range r in range.Cells)
            {
                NeqSimThermoSystem.getThermoSystem().init(0);
                i++;
                for (var k = 0; k < NeqSimThermoSystem.getThermoSystem().getMaxNumberOfPhases(); k++)
                {
                    NeqSimThermoSystem.getThermoSystem().getPhase(k).getComponent((string) r.Text)
                        .setTC(Range["E" + i].Value2 + 273.15);
                    NeqSimThermoSystem.getThermoSystem().getPhase(k).getComponent((string) r.Text)
                        .setPC(Range["F" + i].Value2);
                    NeqSimThermoSystem.getThermoSystem().getPhase(k).getComponent((string) r.Text)
                        .setAcentricFactor(Range["G" + i].Value2);
                }
            }

            NeqSimThermoSystem.getThermoSystem().init(0);
            NeqSimThermoSystem.getThermoSystem().init(1);
            Range["J7"].Font.Color = ColorTranslator.ToOle(Color.Green);
            button1_Click_1(sender, e);
            Range["J7"].Value2 = "done...updated critical parameters";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ((PhaseEosInterface) NeqSimThermoSystem.getThermoSystem().getPhase(0)).displayInteractionCoefficients("");
        }

        private void activteSHeet()
        {
            var rangeClear = Range["A2", "H100"];
           //rangeClear.Clear();
            rangeClear.Font.Color = ColorTranslator.ToOle(Color.DarkGray);

            var rangeBLack = Range["B2", "B200"];
            rangeBLack.Font.Color = ColorTranslator.ToOle(Color.Black);

            var a = NeqSimThermoSystem.getThermoSystem().getPhase(0).getNumberOfComponents() + 1;
            var range = Range["B2", "B" + a];
            var extrange = Range["B" + (a+1), "B200"];
            var i = 0;
            string number;
            foreach (Range r in range.Cells)
              //  if (string.IsNullOrEmpty((string) r.Text))
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
                    Range["E" + number].Value2 =
                        NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getTC() - 273.15;
                    Range["F" + number].Value2 =
                        NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i).getPC();
                    Range["G" + number].Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(i)
                        .getAcentricFactor();
                    i++;
                }

            
            int numb2 = i;
            foreach (Range r in extrange.Cells)
                if (!string.IsNullOrEmpty((string)r.Text))
                {
                    number = (i + 2).ToString();
                    Range["A" + number].Value2 = "";
                    Range["B" + number].Value2 = "";
                    Range["C" + number].Value2 = "";
                    Range["D" + number].Value2 = "";
                    Range["E" + number].Value2 = "";
                    Range["F" + number].Value2 = "";
                    Range["G" + number].Value2 = "";
                    i++;
                }
                else break;
            
            number = (numb2 + 3).ToString();
            Range["A" + number].Value2 = "Total %";
            Range["B" + number].Value2 = "=sum(B2:B" + (numb2 + 1) + ")";
        }
    }
}