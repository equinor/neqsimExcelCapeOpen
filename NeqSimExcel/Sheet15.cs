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
using PVTsimulation.simulation;

namespace NeqSimExcel
{
    public partial class Sheet15
    {
        private void Sheet15_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet15_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.calcButton.Click += new System.EventHandler(this.calcButton_Click);
            this.PVTcalcCombobox.SelectedIndexChanged += new System.EventHandler(this.PVTcalcCombobox_SelectedIndexChanged);
            this.calcCompBUtton.Click += new System.EventHandler(this.calcCompBUtton_Click);
            this.Startup += new System.EventHandler(this.Sheet15_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet15_Shutdown);

        }

        #endregion

        private void PVTcalcCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Excel.Range rangeClear = this.Range["C6", "O100"];
            rangeClear.Clear();

            Excel.Range experimentalTitle = this.Range["C5"];
            Excel.Range calc1RowTitle = this.Range["C6"];
            Excel.Range calc2RowTitle = this.Range["D6"];
            Excel.Range calc3RowTitle = this.Range["E6"];
            Excel.Range calc4RowTitle = this.Range["F6"];
            Excel.Range calc5RowTitle = this.Range["G6"];
            Excel.Range calc6RowTitle = this.Range["H6"];

            Excel.Range calculatedTitle = this.Range["I5"];
            Excel.Range calc7RowTitle = this.Range["I6"];
            Excel.Range calc8RowTitle = this.Range["J6"];
            Excel.Range calc9RowTitle = this.Range["K6"];
            Excel.Range calc10RowTitle = this.Range["L6"];


            experimentalTitle.Value2 = "Experimental data";

            // PVTcalcCombobox.
            if (PVTcalcCombobox.SelectedItem == "Wax content")
            {
                this.Range["C6"].Value2 = "Wax content [wt%]";
            }
            else if (PVTcalcCombobox.SelectedItem == "CME")
            {
                calc1RowTitle.Value2 = "relative volume (V/Vsat)";
                calc2RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc3RowTitle.Value2 = "Z-gas";
                calc4RowTitle.Value2 = "Y-factor";
                calc5RowTitle.Value2 = "IsoThermalCompressibility (1/bar)";
            }
            else if (PVTcalcCombobox.SelectedItem == "GOR" || PVTcalcCombobox.SelectedItem == "Separator test")
            {
                calc1RowTitle.Value2 = "GOR [Sm3/Sm3]";
                calc2RowTitle.Value2 = "Bo-factor";
            }
            else if (PVTcalcCombobox.SelectedItem == "CVD")
            {
                calc1RowTitle.Value2 = "relativeVolume";
                calc2RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc3RowTitle.Value2 = "Z-gas";
                calc4RowTitle.Value2 = "Z-mix";
                calc5RowTitle.Value2 = "cummulative mole% depleted";
            }
            else if (PVTcalcCombobox.SelectedItem == "Differential liberation")
            {
                calc1RowTitle.Value2 = "Bo";
                calc2RowTitle.Value2 = "Bg";
                calc3RowTitle.Value2 = "Rs";
                calc4RowTitle.Value2 = "gas gravity";
                calc5RowTitle.Value2 = "Zgas";
                calc6RowTitle.Value2 = "gas standard volume";
            }

          
        }

        private void calcButton_Click(object sender, EventArgs e)
        {
            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();

            Excel.Range column1range = this.Range["A7", "A100"];
            Excel.Range column2range = this.Range["B7", "B100"];


            Excel.Range calc1RowTitle = this.Range["C6"];
            Excel.Range calc2RowTitle = this.Range["D6"];
            Excel.Range calc3RowTitle = this.Range["E6"];
            Excel.Range calc4RowTitle = this.Range["F6"];
            Excel.Range calc5RowTitle = this.Range["G6"];
            Excel.Range calc6RowTitle = this.Range["H6"];
            Excel.Range calc7RowTitle = this.Range["I6"];


            int number = 0;
            foreach (Excel.Range r in column1range.Cells)
            {
                string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    number++;
                }
            }

            double[] temperatures = new double[number];
            double[] pressures = new double[number];
            double[][] expData = new double[1][];
            expData[0] = new double[number];
            number = 0;
            foreach (Excel.Range r in column1range.Cells)
            {
                string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    temperatures[number] = r.Value2 + 273.15;
                    pressures[number] = this.Range["B" + (number + 7)].Value2;
                    expData[0][number] = this.Range["C" + (number + 7)].Value2; ;
                    number++;
                }
            }

            if (PVTcalcCombobox.SelectedItem == "Wax content")
            {
                calc2RowTitle.Value2 = "calculated wax content [wt%]";
                calc3RowTitle.Value2 = "deviation [%]";
                calc4RowTitle.Value2 = "";
                calc5RowTitle.Value2 = "";

                WaxFractionSim cmdSim = new WaxFractionSim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.setExperimentalData(expData);

                cmdSim.runTuning();
                double[] parameters = thermoSystem.getWaxModel().getWaxParameters();
               // cmdSim.runCalc();


                cmdSim = new WaxFractionSim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.runCalc();
               
                number = 0;
                foreach (double val in cmdSim.getWaxFraction())
                {
                    this.Range["D" + (number + 7)].Value2 = val * 100.0;
                    this.Range["E" + (number + 7)].Value2 = ((val * 100.0 - this.Range["C" + (number + 7)].Value2)/this.Range["C" + (number + 7)].Value2)*100.0;
                    number++;
                }

                this.Range["F1"].Value2 = parameters[0];
                this.Range["F2"].Value2 = parameters[1];
                this.Range["F3"].Value2 = parameters[2];
            }
            if (PVTcalcCombobox.SelectedItem == "CME")
            {
                calc2RowTitle.Value2 = "relative volume (V/Vsat)";
                calc3RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc4RowTitle.Value2 = "Z-gas";
                calc5RowTitle.Value2 = "Y-factor";
                calc6RowTitle.Value2 = "IsoThermalCompressibility (1/bar)";
                thermoSystem.setHeavyTBPfractionAsPlusFraction();
                ConstantMassExpansion cmeSim = new ConstantMassExpansion(thermoSystem);
                cmeSim.setTemperaturesAndPressures(temperatures, pressures);
                cmeSim.setExperimentalData(expData);
                cmeSim.getOptimizer().setMaxNumberOfIterations(5);
                cmeSim.runTuning();
              //  double[] parameters = {thermoSystem.getCharacterization().getPlusFractionModel().getMPlus()};
                // cmdSim.runCalc();


               // cmeSim = new ConstantMassExpansion(thermoSystem);
              //  cmeSim.setTemperature(temperatures[0]);
              ////  cmeSim.setTemperaturesAndPressures(temperatures, pressures);
             //   cmeSim.runCalc();

                number = 0;
                foreach (double val in cmeSim.getRelativeVolume())
                {
                    this.Range["I" + (number + 7)].Value2 = val;
                   // this.Range["E" + (number + 7)].Value2 = ((val * 100.0 - this.Range["C" + (number + 7)].Value2) / this.Range["C" + (number + 7)].Value2) * 100.0;
                    number++;
                }

              //  this.Range["F1"].Value2 = parameters[0];
            }

         //   double[] parameters = thermoSystem.getWaxModel().getWaxParameters();


          
        }

        private void calcCompBUtton_Click(object sender, EventArgs e)
        {
            Excel.Range rangeClear = this.Range["I1", "O100"];
            rangeClear.Clear();

            SystemInterface thermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem().clone();

            Excel.Range column1range = this.Range["A7", "A100"];
            Excel.Range column2range = this.Range["B7", "B100"];

            Excel.Range calc1RowTitle = this.Range["C6"];
            Excel.Range calc2RowTitle = this.Range["D6"];
            Excel.Range calc3RowTitle = this.Range["E6"];
            Excel.Range calc4RowTitle = this.Range["F6"];
            Excel.Range calc5RowTitle = this.Range["G6"];
            Excel.Range calc6RowTitle = this.Range["H6"];
            Excel.Range calc7RowTitle = this.Range["I6"];
            Excel.Range calculatedTitle = this.Range["I5"];
            Excel.Range calc8RowTitle = this.Range["J6"];
            Excel.Range calc9RowTitle = this.Range["K6"];
            Excel.Range calc10RowTitle = this.Range["L6"];
            Excel.Range calc11RowTitle = this.Range["M6"];
            Excel.Range calc12RowTitle = this.Range["N6"];

            int number = 0;

            calculatedTitle.Value2 = "Calculated";
            
            foreach (Excel.Range r in column1range.Cells)
            {
                string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    number++;
                }
            }

            double[] temperatures = new double[number];
            double[] pressures = new double[number];

            number = 0;
            foreach (Excel.Range r in column1range.Cells)
            {
                string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    temperatures[number] = r.Value2 + 273.15;
                    pressures[number] = this.Range["B" + (number + 7)].Value2;
                    number++;
                }
            }


            if (PVTcalcCombobox.SelectedItem == "GOR" || PVTcalcCombobox.SelectedItem == "Separator test")
            {
                calc7RowTitle.Value2 = "GOR [Sm3/Sm3]";
                calc8RowTitle.Value2 = "Bo-factor";
                number = 0;

                if (PVTcalcCombobox.SelectedItem == "Separator test")
                {
                    SeparatorTest sepSim = new SeparatorTest(thermoSystem);
                    sepSim.setSeparatorConditions(temperatures, pressures);
                    sepSim.runCalc();
                    foreach (double val in sepSim.getGOR())
                    {
                        this.Range["I" + (number + 7)].Value2 = val;
                        this.Range["J" + (number + 7)].Value2 = sepSim.getBofactor()[number];
                        number++;
                    }
                }
                else
                {
                    GOR sepSim = new GOR(thermoSystem);
                    sepSim.setTemperaturesAndPressures(temperatures, pressures);
                    sepSim.runCalc();
                    foreach (double val in sepSim.getGOR())
                    {
                        this.Range["I" + (number + 7)].Value2 = val;
                        this.Range["J" + (number + 7)].Value2 = sepSim.getBofactor()[number];
                        number++;
                    }
                }

            }
            else if (PVTcalcCombobox.SelectedItem == "CME")
            {
                calc7RowTitle.Value2 = "relative volume (V/Vsat)";
                calc8RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc9RowTitle.Value2 = "Z-gas";
                calc10RowTitle.Value2 = "Y-factor";
                calc11RowTitle.Value2 = "IsoThermalCompressibility (1/bar)";
                ConstantMassExpansion cmeSim = new ConstantMassExpansion(thermoSystem);
                cmeSim.setTemperature(temperatures[0]);
                cmeSim.setPressures(pressures);
                cmeSim.runCalc();
                number = 0;
                foreach (double val in cmeSim.getRelativeVolume())
                {
                    this.Range["I" + (number + 7)].Value2 = val;
                    if (cmeSim.getLiquidRelativeVolume()[number] > 1e-50) this.Range["J" + (number + 7)].Value2 = cmeSim.getLiquidRelativeVolume()[number];
                    if (cmeSim.getZgas()[number] > 1e-20) this.Range["K" + (number + 7)].Value2 = cmeSim.getZgas()[number];
                    if (cmeSim.getYfactor()[number] > 1e-20) this.Range["L" + (number + 7)].Value2 = cmeSim.getYfactor()[number];
                    if (cmeSim.getIsoThermalCompressibility()[number] > 1e-20) this.Range["M" + (number + 7)].Value2 = cmeSim.getIsoThermalCompressibility()[number];
                    number++;
                }
                this.Range["A" + (number + 7)].Value2 = temperatures[0] - 273.15;
                this.Range["B" + (number + 7)].Value2 = cmeSim.getSaturationPressure();
                this.Range["I" + (number + 7)].Value2 = 1.0;
                this.Range["J" + (number + 7)].Value2 = "saturation point";
                this.Range["K" + (number + 7)].Value2 = cmeSim.getZsaturation();
                this.Range["M" + (number + 7)].Value2 = cmeSim.getSaturationIsoThermalCompressibility();

            }
            else if (PVTcalcCombobox.SelectedItem == "CVD")
            {
                calc7RowTitle.Value2 = "relativeVolume";
                calc8RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc9RowTitle.Value2 = "Z-gas";
                calc10RowTitle.Value2 = "Z-mix";
                calc11RowTitle.Value2 = "cummulative mole% depleted";

                ConstantVolumeDepletion cmdSim = new ConstantVolumeDepletion(thermoSystem);
                cmdSim.setTemperature(temperatures[0]);
                cmdSim.setPressures(pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (double val in cmdSim.getRelativeVolume())
                {
                    this.Range["I" + (number + 7)].Value2 = val;
                    this.Range["J" + (number + 7)].Value2 = cmdSim.getLiquidRelativeVolume()[number];
                    this.Range["K" + (number + 7)].Value2 = cmdSim.getZgas()[number];
                    this.Range["L" + (number + 7)].Value2 = cmdSim.getZmix()[number];
                    this.Range["M" + (number + 7)].Value2 = cmdSim.getCummulativeMolePercDepleted()[number];
                    number++;
                }
                this.Range["A" + (number + 7)].Value2 = temperatures[0] - 273.15;
                this.Range["B" + (number + 7)].Value2 = cmdSim.getSaturationPressure();
                this.Range["I" + (number + 7)].Value2 = 1.0;
                this.Range["K" + (number + 7)].Value2 = cmdSim.getZsaturation();
                this.Range["L" + (number + 7)].Value2 = cmdSim.getZsaturation();
                this.Range["M" + (number + 7)].Value2 = "saturation point";
            }

            else if (PVTcalcCombobox.SelectedItem == "Differential liberation")
            {
                calc7RowTitle.Value2 = "Bo";
                calc8RowTitle.Value2 = "Bg";
                calc9RowTitle.Value2 = "Rs";
                calc10RowTitle.Value2 = "gas gravity";
                calc11RowTitle.Value2 = "Zgas";
                calc12RowTitle.Value2 = "gas standard volume";

                DifferentialLiberation difLibSIm = new DifferentialLiberation(thermoSystem);
                difLibSIm.setTemperature(temperatures[0]);
                difLibSIm.setPressures(pressures);
                difLibSIm.runCalc();
                number = 0;
                foreach (double val in difLibSIm.getBo())
                {
                    this.Range["I" + (number + 7)].Value2 = val;
                    this.Range["J" + (number + 7)].Value2 = difLibSIm.getBg()[(number)];
                    this.Range["K" + (number + 7)].Value2 = difLibSIm.getRs()[(number)];
                    this.Range["L" + (number + 7)].Value2 = difLibSIm.getRelGasGravity()[(number)];
                    this.Range["M" + (number + 7)].Value2 = difLibSIm.getZgas()[(number)];
                    this.Range["N" + (number + 7)].Value2 = difLibSIm.getGasStandardVolume()[(number)];
                    number++;
                }
            }
            else if (PVTcalcCombobox.SelectedItem == "Wax content")
            {
                calc7RowTitle.Value2 = "wt% wax";
                calc8RowTitle.Value2 = "";
                calc9RowTitle.Value2 = "";
                calc10RowTitle.Value2 = "";
                calc11RowTitle.Value2 = "";

                WaxFractionSim cmdSim = new WaxFractionSim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (double val in cmdSim.getWaxFraction())
                {
                    this.Range["I" + (number + 7)].Value2 = val * 100.0;
                    number++;
                }
            }
        }
    }
}
