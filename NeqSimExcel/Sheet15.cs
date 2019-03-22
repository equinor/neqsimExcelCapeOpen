using System;
using Microsoft.Office.Interop.Excel;
using neqsim.PVTsimulation.simulation;
using neqsim.thermo.system;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet15
    {
        private void Sheet15_Startup(object sender, EventArgs e)
        {
            PVTcalcCombobox.SelectedIndex = 0;
            selectFluidCombobox.Visible = false;

        }

        private void Sheet15_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.calcButton.Click += new System.EventHandler(this.calcButton_Click);
            this.PVTcalcCombobox.SelectedIndexChanged += new System.EventHandler(this.PVTcalcCombobox_SelectedIndexChanged);
            this.calcCompBUtton.Click += new System.EventHandler(this.calcCompBUtton_Click);
            this.selectFluidCombobox.SelectedIndexChanged += new System.EventHandler(this.PVTcalcCombobox_SelectedIndexChanged);
            this.Startup += new System.EventHandler(this.Sheet15_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet15_Shutdown);

        }

        #endregion

        private void PVTcalcCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var rangeClear = Range["C6", "O100"];
            rangeClear.Clear();

            var experimentalTitle = Range["C5"];
            var calc1RowTitle = Range["C6"];
            var calc2RowTitle = Range["D6"];
            var calc3RowTitle = Range["E6"];
            var calc4RowTitle = Range["F6"];
            var calc5RowTitle = Range["G6"];
            var calc6RowTitle = Range["H6"];

            var calculatedTitle = Range["I5"];
            var calc7RowTitle = Range["I6"];
            var calc8RowTitle = Range["J6"];
            var calc9RowTitle = Range["K6"];
            var calc10RowTitle = Range["L6"];


            experimentalTitle.Value2 = "Experimental data";

            // PVTcalcCombobox.
            if (PVTcalcCombobox.SelectedItem.ToString().Equals("Wax content"))
            {
                Range["C6"].Value2 = "Wax content [wt%]";
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("CME"))
            {
                calc1RowTitle.Value2 = "relative volume (V/Vsat)";
                calc2RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc3RowTitle.Value2 = "Z-gas";
                calc4RowTitle.Value2 = "Y-factor";
                calc5RowTitle.Value2 = "IsoThermalCompressibility (1/bar)";
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("GOR") || PVTcalcCombobox.SelectedItem.ToString().Equals("Separator test"))
            {
                calc1RowTitle.Value2 = "GOR [Sm3/Sm3]";
                calc2RowTitle.Value2 = "Bo-factor";
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("CVD"))
            {
                calc1RowTitle.Value2 = "relativeVolume";
                calc2RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc3RowTitle.Value2 = "Z-gas";
                calc4RowTitle.Value2 = "Z-mix";
                calc5RowTitle.Value2 = "cummulative mole% depleted";
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("Differential liberation"))
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
            var thermoSystem = NeqSimThermoSystem.getThermoSystem();

            var column1range = Range["A7", "A100"];
            var column2range = Range["B7", "B100"];


            var calc1RowTitle = Range["C6"];
            var calc2RowTitle = Range["D6"];
            var calc3RowTitle = Range["E6"];
            var calc4RowTitle = Range["F6"];
            var calc5RowTitle = Range["G6"];
            var calc6RowTitle = Range["H6"];
            var calc7RowTitle = Range["I6"];


            var number = 0;
            foreach (Range r in column1range.Cells)
            {
                var text = (string) r.Text;
                if (!string.IsNullOrEmpty(text)) number++;
            }

            var temperatures = new double[number];
            var pressures = new double[number];
            var expData = new double[1][];
            expData[0] = new double[number];
            number = 0;
            foreach (Range r in column1range.Cells)
            {
                var text = (string) r.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    temperatures[number] = r.Value2 + 273.15;
                    pressures[number] = Range["B" + (number + 7)].Value2;
                    expData[0][number] = Range["C" + (number + 7)].Value2;
                    ;
                    number++;
                }
            }

            if (PVTcalcCombobox.SelectedItem.ToString().Equals("Wax content"))
            {
                calc2RowTitle.Value2 = "calculated wax content [wt%]";
                calc3RowTitle.Value2 = "deviation [%]";
                calc4RowTitle.Value2 = "";
                calc5RowTitle.Value2 = "";

                var cmdSim = new WaxFractionSim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.setExperimentalData(expData);

                cmdSim.runTuning();
                var parameters = thermoSystem.getWaxModel().getWaxParameters();
                // cmdSim.runCalc();


                cmdSim = new WaxFractionSim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.runCalc();

                number = 0;
                foreach (var val in cmdSim.getWaxFraction())
                {
                    Range["D" + (number + 7)].Value2 = val * 100.0;
                    Range["E" + (number + 7)].Value2 = (val * 100.0 - Range["C" + (number + 7)].Value2) /
                                                       Range["C" + (number + 7)].Value2 * 100.0;
                    number++;
                }

                Range["F1"].Value2 = parameters[0];
                Range["F2"].Value2 = parameters[1];
                Range["F3"].Value2 = parameters[2];
            }

            if (PVTcalcCombobox.SelectedItem.ToString().Equals("CME"))
            {
                calc2RowTitle.Value2 = "relative volume (V/Vsat)";
                calc3RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc4RowTitle.Value2 = "Z-gas";
                calc5RowTitle.Value2 = "Y-factor";
                calc6RowTitle.Value2 = "IsoThermalCompressibility (1/bar)";
                thermoSystem.setHeavyTBPfractionAsPlusFraction();
                var cmeSim = new ConstantMassExpansion(thermoSystem);
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
                foreach (var val in cmeSim.getRelativeVolume())
                {
                    Range["I" + (number + 7)].Value2 = val;
                    // this.Range["E" + (number + 7)].Value2 = ((val * 100.0 - this.Range["C" + (number + 7)].Value2) / this.Range["C" + (number + 7)].Value2) * 100.0;
                    number++;
                }

                //  this.Range["F1"].Value2 = parameters[0];
            }

            //   double[] parameters = thermoSystem.getWaxModel().getWaxParameters();
        }

        private void calcCompBUtton_Click(object sender, EventArgs e)
        {
            var rangeClear = Range["I1", "O100"];
            rangeClear.Clear();

            var thermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem().clone();

            var column1range = Range["A7", "A100"];
            var column2range = Range["B7", "B100"];

            var calc1RowTitle = Range["C6"];
            var calc2RowTitle = Range["D6"];
            var calc3RowTitle = Range["E6"];
            var calc4RowTitle = Range["F6"];
            var calc5RowTitle = Range["G6"];
            var calc6RowTitle = Range["H6"];
            var calc7RowTitle = Range["I6"];
            var calculatedTitle = Range["I5"];
            var calc8RowTitle = Range["J6"];
            var calc9RowTitle = Range["K6"];
            var calc10RowTitle = Range["L6"];
            var calc11RowTitle = Range["M6"];
            var calc12RowTitle = Range["N6"];

            var number = 0;

            calculatedTitle.Value2 = "Calculated";

            foreach (Range r in column1range.Cells)
            {
                var text = (string)r.Text;
                if (!string.IsNullOrEmpty(text)) number++;
            }

            var temperatures = new double[number];
            var pressures = new double[number];

            number = 0;
            foreach (Range r in column1range.Cells)
            {
                var text = (string)r.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    temperatures[number] = r.Value2 + 273.15;
                    pressures[number] = Range["B" + (number + 7)].Value2;
                    number++;
                }
            }


            if (PVTcalcCombobox.SelectedItem.ToString().Equals("GOR") || PVTcalcCombobox.SelectedItem.ToString().Equals("Separator test"))
            {
                calc7RowTitle.Value2 = "GOR [Sm3/Sm3]";
                calc8RowTitle.Value2 = "Bo-factor";
                number = 0;

                if (PVTcalcCombobox.SelectedItem.ToString().Equals("Separator test"))
                {
                    var sepSim = new SeparatorTest(thermoSystem);
                    sepSim.setSeparatorConditions(temperatures, pressures);
                    sepSim.runCalc();
                    foreach (var val in sepSim.getGOR())
                    {
                        Range["I" + (number + 7)].Value2 = val;
                        Range["J" + (number + 7)].Value2 = sepSim.getBofactor()[number];
                        number++;
                    }
                }
                else
                {
                    var sepSim = new GOR(thermoSystem);
                    sepSim.setTemperaturesAndPressures(temperatures, pressures);
                    sepSim.runCalc();
                    foreach (var val in sepSim.getGOR())
                    {
                        Range["I" + (number + 7)].Value2 = val;
                        Range["J" + (number + 7)].Value2 = sepSim.getBofactor()[number];
                        number++;
                    }
                }
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("CME"))
            {
                calc7RowTitle.Value2 = "relative volume (V/Vsat)";
                calc8RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc9RowTitle.Value2 = "Z-gas";
                calc10RowTitle.Value2 = "Y-factor";
                calc11RowTitle.Value2 = "IsoThermalCompressibility (1/bar)";
                var cmeSim = new ConstantMassExpansion(thermoSystem);
                cmeSim.setTemperature(temperatures[0]);
                cmeSim.setPressures(pressures);
                cmeSim.runCalc();
                number = 0;
                foreach (var val in cmeSim.getRelativeVolume())
                {
                    Range["I" + (number + 7)].Value2 = val;
                    if (cmeSim.getLiquidRelativeVolume()[number] > 1e-50)
                        Range["J" + (number + 7)].Value2 = cmeSim.getLiquidRelativeVolume()[number];
                    if (cmeSim.getZgas()[number] > 1e-20) Range["K" + (number + 7)].Value2 = cmeSim.getZgas()[number];
                    if (cmeSim.getYfactor()[number] > 1e-20)
                        Range["L" + (number + 7)].Value2 = cmeSim.getYfactor()[number];
                    if (cmeSim.getIsoThermalCompressibility()[number] > 1e-20)
                        Range["M" + (number + 7)].Value2 = cmeSim.getIsoThermalCompressibility()[number];
                    number++;
                }

                Range["A" + (number + 7)].Value2 = temperatures[0] - 273.15;
                Range["B" + (number + 7)].Value2 = cmeSim.getSaturationPressure();
                Range["I" + (number + 7)].Value2 = 1.0;
                Range["J" + (number + 7)].Value2 = "saturation point";
                Range["K" + (number + 7)].Value2 = cmeSim.getZsaturation();
                Range["M" + (number + 7)].Value2 = cmeSim.getSaturationIsoThermalCompressibility();
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("CVD"))
            {
                calc7RowTitle.Value2 = "relativeVolume";
                calc8RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc9RowTitle.Value2 = "Z-gas";
                calc10RowTitle.Value2 = "Z-mix";
                calc11RowTitle.Value2 = "cummulative mole% depleted";

                var cmdSim = new ConstantVolumeDepletion(thermoSystem);
                cmdSim.setTemperature(temperatures[0]);
                cmdSim.setPressures(pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (var val in cmdSim.getRelativeVolume())
                {
                    Range["I" + (number + 7)].Value2 = val;
                    Range["J" + (number + 7)].Value2 = cmdSim.getLiquidRelativeVolume()[number];
                    Range["K" + (number + 7)].Value2 = cmdSim.getZgas()[number];
                    Range["L" + (number + 7)].Value2 = cmdSim.getZmix()[number];
                    Range["M" + (number + 7)].Value2 = cmdSim.getCummulativeMolePercDepleted()[number];
                    number++;
                }

                Range["A" + (number + 7)].Value2 = temperatures[0] - 273.15;
                Range["B" + (number + 7)].Value2 = cmdSim.getSaturationPressure();
                Range["I" + (number + 7)].Value2 = 1.0;
                Range["K" + (number + 7)].Value2 = cmdSim.getZsaturation();
                Range["L" + (number + 7)].Value2 = cmdSim.getZsaturation();
                Range["M" + (number + 7)].Value2 = "saturation point";
            }

            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("Differential liberation"))
            {
                calc7RowTitle.Value2 = "Bo";
                calc8RowTitle.Value2 = "Bg";
                calc9RowTitle.Value2 = "Rs";
                calc10RowTitle.Value2 = "gas gravity";
                calc11RowTitle.Value2 = "Zgas";
                calc12RowTitle.Value2 = "gas standard volume";

                var difLibSIm = new DifferentialLiberation(thermoSystem);
                difLibSIm.setTemperature(temperatures[0]);
                difLibSIm.setPressures(pressures);
                difLibSIm.runCalc();
                number = 0;
                foreach (var val in difLibSIm.getBo())
                {
                    Range["I" + (number + 7)].Value2 = val;
                    Range["J" + (number + 7)].Value2 = difLibSIm.getBg()[number];
                    Range["K" + (number + 7)].Value2 = difLibSIm.getRs()[number];
                    Range["L" + (number + 7)].Value2 = difLibSIm.getRelGasGravity()[number];
                    Range["M" + (number + 7)].Value2 = difLibSIm.getZgas()[number];
                    Range["N" + (number + 7)].Value2 = difLibSIm.getGasStandardVolume()[number];
                    number++;
                }
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("Wax content"))
            {
                calc7RowTitle.Value2 = "wt% wax";
                calc8RowTitle.Value2 = "";
                calc9RowTitle.Value2 = "";
                calc10RowTitle.Value2 = "";
                calc11RowTitle.Value2 = "";

                var cmdSim = new WaxFractionSim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (var val in cmdSim.getWaxFraction())
                {
                    Range["I" + (number + 7)].Value2 = val * 100.0;
                    number++;
                }
            }

        }
    }
}