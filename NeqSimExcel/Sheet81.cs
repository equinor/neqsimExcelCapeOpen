using System;
using System.Collections.Generic;
using System.Security.Principal;
using Microsoft.Office.Interop.Excel;
using neqsim.PVTsimulation.simulation;
using neqsim.thermo.system;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class PVTSheet
    {
        private void Sheet8_Startup(object sender, EventArgs e)
        {
            // selectFluidCombobox.Hide();


        }

        private void Sheet8_Shutdown(object sender, EventArgs e)
        {
        }

        private void ActivateWorkSheet()
        {
            selectFluidCombobox.Items.Clear();
            try
            {
                //               NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter test = new NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter test = new NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfo1TableAdapter test = new neqsimdatabaseDataSetTableAdapters.fluidinfo1TableAdapter();

                var userName = WindowsIdentity.GetCurrent().Name;
                userName = userName.Replace("STATOIL-NET\\", "");
                userName = userName.Replace("WIN-NTNU-NO\\", "");
                userName = userName.ToLower();

               // var tt = test.GetDataBy(userName);
                //               NeqSimExcel.DataSet1.fluidinfoDataTable tt = test.GetData(userName);
                var names = new List<string>();
                //names.Add("CPApackage");
                //names.Add(WindowsIdentity.GetCurrent().Name);
             

                //   packageNames = names.ToArray();
                //   fluidListNameComboBox.Items.Add(names.ToList());
                //selectFluidCombobox.SelectedIndex = 0;
            }
            catch (Exception excet)
            {
                Console.WriteLine("Error " + excet.Message);
            }

            PVTcalcCombobox.SelectedIndex = 0;
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
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.ActivateWorkSheet);
            this.Startup += new System.EventHandler(this.Sheet8_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet8_Shutdown);

        }

        #endregion

        private void calcButton_Click(object sender, EventArgs e)
        {
            var rangeClear = Range["D1", "L100"];
            rangeClear.Clear();

            var thermoSystem = (SystemInterface) NeqSimThermoSystem.getThermoSystem().clone();

            var column1range = Range["A7", "A100"];
            var column2range = Range["B7", "B100"];
            var column3range = Range["C7", "C100"];

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

            number = 0;
            foreach (Range r in column1range.Cells)
            {
                var text = (string) r.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    temperatures[number] = r.Value2 + 273.15;
                    pressures[number] = Range["B" + (number + 7)].Value2;
                    number++;
                }
            }


            if (PVTcalcCombobox.SelectedItem.ToString().Equals("GOR") ||
                PVTcalcCombobox.SelectedItem.ToString().Equals("Separator test"))
            {
                calc1RowTitle.Value2 = "GOR [Sm3/Sm3]";
                calc2RowTitle.Value2 = "Bo-factor";
                number = 0;

                if (PVTcalcCombobox.SelectedItem.ToString().Equals("Separator test"))
                {
                    var sepSim = new SeparatorTest(thermoSystem);
                    sepSim.setSeparatorConditions(temperatures, pressures);
                    sepSim.runCalc();
                    foreach (var val in sepSim.getGOR())
                    {
                        Range["C" + (number + 7)].Value2 = val;
                        Range["D" + (number + 7)].Value2 = sepSim.getBofactor()[number];
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
                        Range["C" + (number + 7)].Value2 = val;
                        Range["D" + (number + 7)].Value2 = sepSim.getBofactor()[number];
                        number++;
                    }
                }
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("CME"))
            {
                calc1RowTitle.Value2 = "relative volume (V/Vsat)";
                calc2RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc3RowTitle.Value2 = "Z-gas";
                calc4RowTitle.Value2 = "Y-factor";
                calc5RowTitle.Value2 = "IsoThermalCompressibility (1/bar)";
                var cmeSim = new ConstantMassExpansion(thermoSystem);
                cmeSim.setTemperature(temperatures[0]);
                cmeSim.setPressures(pressures);
                cmeSim.runCalc();
                number = 0;
                foreach (var val in cmeSim.getRelativeVolume())
                {
                    Range["C" + (number + 7)].Value2 = val;
                    if (cmeSim.getLiquidRelativeVolume()[number] > 1e-50)
                        Range["D" + (number + 7)].Value2 = cmeSim.getLiquidRelativeVolume()[number];
                    if (cmeSim.getZgas()[number] > 1e-20) Range["E" + (number + 7)].Value2 = cmeSim.getZgas()[number];
                    if (cmeSim.getYfactor()[number] > 1e-20)
                        Range["F" + (number + 7)].Value2 = cmeSim.getYfactor()[number];
                    if (cmeSim.getIsoThermalCompressibility()[number] > 1e-20)
                        Range["G" + (number + 7)].Value2 = cmeSim.getIsoThermalCompressibility()[number];
                    number++;
                }

                Range["A" + (number + 7)].Value2 = temperatures[0] - 273.15;
                Range["B" + (number + 7)].Value2 = cmeSim.getSaturationPressure();
                Range["C" + (number + 7)].Value2 = 1.0;
                Range["D" + (number + 7)].Value2 = "saturation point";
                Range["E" + (number + 7)].Value2 = cmeSim.getZsaturation();
                Range["G" + (number + 7)].Value2 = cmeSim.getSaturationIsoThermalCompressibility();
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("CVD"))
            {
                calc1RowTitle.Value2 = "relativeVolume";
                calc2RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc3RowTitle.Value2 = "Z-gas";
                calc4RowTitle.Value2 = "Z-mix";
                calc5RowTitle.Value2 = "cummulative mole% depleted";

                var cmdSim = new ConstantVolumeDepletion(thermoSystem);
                cmdSim.setTemperature(temperatures[0]);
                cmdSim.setPressures(pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (var val in cmdSim.getRelativeVolume())
                {
                    Range["C" + (number + 7)].Value2 = val;
                    Range["D" + (number + 7)].Value2 = cmdSim.getLiquidRelativeVolume()[number];
                    Range["E" + (number + 7)].Value2 = cmdSim.getZgas()[number];
                    Range["F" + (number + 7)].Value2 = cmdSim.getZmix()[number];
                    Range["G" + (number + 7)].Value2 = cmdSim.getCummulativeMolePercDepleted()[number];
                    number++;
                }

                Range["A" + (number + 7)].Value2 = temperatures[0] - 273.15;
                Range["B" + (number + 7)].Value2 = cmdSim.getSaturationPressure();
                Range["C" + (number + 7)].Value2 = 1.0;
                Range["E" + (number + 7)].Value2 = cmdSim.getZsaturation();
                Range["F" + (number + 7)].Value2 = cmdSim.getZsaturation();
                Range["G" + (number + 7)].Value2 = "saturation point";
            }

            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("Differential liberation"))
            {
                calc1RowTitle.Value2 = "Bo";
                calc2RowTitle.Value2 = "Bg";
                calc3RowTitle.Value2 = "Rs";
                calc4RowTitle.Value2 = "gas gravity";
                calc5RowTitle.Value2 = "Zgas";
                calc6RowTitle.Value2 = "gas standard volume";

                var difLibSIm = new DifferentialLiberation(thermoSystem);
                difLibSIm.setTemperature(temperatures[0]);
                difLibSIm.setPressures(pressures);
                difLibSIm.runCalc();
                number = 0;
                foreach (var val in difLibSIm.getBo())
                {
                    Range["C" + (number + 7)].Value2 = val;
                    Range["D" + (number + 7)].Value2 = difLibSIm.getBg()[number];
                    Range["E" + (number + 7)].Value2 = difLibSIm.getRs()[number];
                    Range["F" + (number + 7)].Value2 = difLibSIm.getRelGasGravity()[number];
                    Range["G" + (number + 7)].Value2 = difLibSIm.getZgas()[number];
                    Range["H" + (number + 7)].Value2 = difLibSIm.getGasStandardVolume()[number];
                    number++;
                }
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("Wax content"))
            {
                calc1RowTitle.Value2 = "wt% wax";
                calc2RowTitle.Value2 = "";
                calc3RowTitle.Value2 = "";
                calc4RowTitle.Value2 = "";
                calc5RowTitle.Value2 = "";

                var cmdSim = new WaxFractionSim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (var val in cmdSim.getWaxFraction())
                {
                    Range["C" + (number + 7)].Value2 = val * 100.0;
                    number++;
                }
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("Viscosity"))
            {
                calc1RowTitle.Value2 = "Gas viscosity [kg/m*sec]";
                calc2RowTitle.Value2 = "Oil viscosity [kg/m*sec]";
                calc3RowTitle.Value2 = "Aqueous viscosity [kg/m*sec]";
                calc4RowTitle.Value2 = "";
                calc5RowTitle.Value2 = "";

                var cmdSim = new ViscositySim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (var val in cmdSim.getGasViscosity())
                {
                    Range["C" + (number + 7)].Value2 = val;
                    Range["D" + (number + 7)].Value2 = cmdSim.getOilViscosity()[number];
                    Range["E" + (number + 7)].Value2 = cmdSim.getAqueousViscosity()[number];
                    number++;
                }
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("Viscosity oil/wax"))
            {
                calc1RowTitle.Value2 = "share rate [1/sec]";
                calc2RowTitle.Value2 = "wt% wax";
                calc3RowTitle.Value2 = "Oil viscosity [kg/m*sec]";
                calc4RowTitle.Value2 = "Oil/wax viscosity [kg/m*sec]";
                calc5RowTitle.Value2 = "";
                calc6RowTitle.Value2 = "";


                var shareRate = new double[number];

                number = 0;
                foreach (Range r in column3range.Cells)
                {
                    var text = (string) r.Text;
                    if (!string.IsNullOrEmpty(text))
                    {
                        shareRate[number] = r.Value2;
                        number++;
                    }
                }


                var cmdSim = new ViscosityWaxOilSim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.setShareRate(shareRate);
                cmdSim.runCalc();
                number = 0;
                foreach (var val in cmdSim.getGasViscosity())
                {
                    Range["D" + (number + 7)].Value2 = cmdSim.getWaxFraction()[number] * 100;
                    Range["E" + (number + 7)].Value2 = cmdSim.getOilViscosity()[number];
                    Range["F" + (number + 7)].Value2 = cmdSim.getOilwaxDispersionViscosity()[number];
                    number++;
                }
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("Swelling test"))
            {
                calc2RowTitle.Value2 = "Relative volume";
                var a2 = Convert.ToInt32(selectFluidCombobox.SelectedItem.ToString());
                var gasInjectionThermoSystem = NeqSimThermoSystem.getThermoSystem().readObject(a2);

                var cmdSim = new SwellingTest((SystemInterface) thermoSystem.clone());
                cmdSim.setInjectionGas(gasInjectionThermoSystem);
                cmdSim.setTemperature(temperatures[0]);
                cmdSim.setCummulativeMolePercentGasInjected(pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (var val in cmdSim.getPressures())
                {
                    Range["C" + (number + 7)].Value2 = cmdSim.getPressures()[number];
                    Range["D" + (number + 7)].Value2 = cmdSim.getRelativeOilVolume()[number];
                    number++;
                }
            }

            rangeClear.Columns.AutoFit();
        }

        private void PVTcalcCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Range["B6"].Value2 = "Pressure [bara]";
            selectFluidCombobox.Visible = false;
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
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("Viscosity oil/wax"))
            {
                calc1RowTitle.Value2 = "Share rate [1/sec]";
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("Swelling test"))
            {
                Range["B6"].Value2 = "GasInjected (fraction of oil)";
                Range["C6"].Value2 = "Pressure [bara]";
                selectFluidCombobox.Visible = true;
                Range["A4"].Value2 = "Injection gas";
                calc2RowTitle.Value2 = "Relative volume";
            }
            else if (PVTcalcCombobox.SelectedItem.ToString().Equals("Slim tube"))
            {
                selectFluidCombobox.Visible = true;
                Range["A4"].Value2 = "Injection gas";
            }

            rangeClear.Columns.AutoFit();
        }
    }
}