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
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Collections;
using System.Security.Principal;
using MySql.Data;

namespace NeqSimExcel
{
    public partial class PVTSheet
    {
        private void Sheet8_Startup(object sender, System.EventArgs e)
        {
          // selectFluidCombobox.Hide();

            
           try
           {
               DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter test = new DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter();
//               NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter test = new NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter();
               //NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter test = new NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter();
               //NeqSimExcelDataSetTableAdapters.fluidinfo1TableAdapter test = new neqsimdatabaseDataSetTableAdapters.fluidinfo1TableAdapter();

               string userName = WindowsIdentity.GetCurrent().Name;
               userName = userName.Replace("STATOIL-NET\\", "");
               userName = userName.Replace("WIN-NTNU-NO\\", "");
               userName = userName.ToLower();

               DatabaseConnection.NeqSimDatabaseSet.fluidinfoDataTable tt = test.GetDataBy(userName);
//               NeqSimExcel.DataSet1.fluidinfoDataTable tt = test.GetData(userName);
               List<string> names = new List<string>();
               //names.Add("CPApackage");
               //names.Add(WindowsIdentity.GetCurrent().Name);
               foreach (DatabaseConnection.NeqSimDatabaseSet.fluidinfoRow row in tt.Rows)
               {
                   names.Add(row.ID.ToString());
                   selectFluidCombobox.Items.Add(row.ID.ToString());
               }
               //   packageNames = names.ToArray();
               //   fluidListNameComboBox.Items.Add(names.ToList());
               //selectFluidCombobox.SelectedIndex = 0;
           }
           catch (Exception excet)
           {
               Console.WriteLine("Error " + excet.Message);
           }
            
        }

        private void Sheet8_Shutdown(object sender, System.EventArgs e)
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
            this.Startup += new System.EventHandler(this.Sheet8_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet8_Shutdown);

        }

        #endregion

        private void calcButton_Click(object sender, EventArgs e)
        {
            Excel.Range rangeClear = this.Range["D1", "L100"];
            rangeClear.Clear();

            SystemInterface thermoSystem = (SystemInterface) NeqSimThermoSystem.getThermoSystem().clone();

            Excel.Range column1range = this.Range["A7", "A100"];
            Excel.Range column2range = this.Range["B7", "B100"];
            Excel.Range column3range = this.Range["C7", "C100"];
            
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

            number = 0;
            foreach (Excel.Range r in column1range.Cells)
            {
                string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    temperatures[number] = r.Value2+273.15;
                    pressures[number] = this.Range["B" + (number + 7)].Value2;
                    number++;
                }
             }


            if (PVTcalcCombobox.SelectedItem.ToString().Equals("GOR") || PVTcalcCombobox.SelectedItem.ToString().Equals("Separator test"))
            {
                calc1RowTitle.Value2 = "GOR [Sm3/Sm3]";
                calc2RowTitle.Value2 = "Bo-factor";
                number = 0;

                if (PVTcalcCombobox.SelectedItem.ToString().Equals("Separator test"))
                {
                    SeparatorTest sepSim = new SeparatorTest(thermoSystem);
                    sepSim.setSeparatorConditions(temperatures, pressures);
                    sepSim.runCalc();
                    foreach (double val in sepSim.getGOR())
                    {
                        this.Range["C" + (number + 7)].Value2 = val;
                        this.Range["D" + (number + 7)].Value2 = sepSim.getBofactor()[number];
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
                        this.Range["C" + (number + 7)].Value2 = val;
                        this.Range["D" + (number + 7)].Value2 = sepSim.getBofactor()[number];
                        number++;
                    }
                }
    
            }
            else if (PVTcalcCombobox.SelectedItem == "CME")
            {
                calc1RowTitle.Value2 = "relative volume (V/Vsat)";
                calc2RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc3RowTitle.Value2 = "Z-gas";
                calc4RowTitle.Value2 = "Y-factor";
                calc5RowTitle.Value2 = "IsoThermalCompressibility (1/bar)";
                ConstantMassExpansion cmeSim = new ConstantMassExpansion(thermoSystem);
                cmeSim.setTemperature(temperatures[0]);
                cmeSim.setPressures(pressures);
                cmeSim.runCalc();
                number = 0;
                foreach (double val in cmeSim.getRelativeVolume())
                {
                    this.Range["C" + (number + 7)].Value2 = val;
                    if (cmeSim.getLiquidRelativeVolume()[number] > 1e-50) this.Range["D" + (number + 7)].Value2 = cmeSim.getLiquidRelativeVolume()[number];
                    if (cmeSim.getZgas()[number] > 1e-20) this.Range["E" + (number + 7)].Value2 = cmeSim.getZgas()[number];
                    if (cmeSim.getYfactor()[number] > 1e-20) this.Range["F" + (number + 7)].Value2 = cmeSim.getYfactor()[number];
                    if (cmeSim.getIsoThermalCompressibility()[number] > 1e-20) this.Range["G" + (number + 7)].Value2 = cmeSim.getIsoThermalCompressibility()[number];
                    number++;
                }
                this.Range["A" + (number + 7)].Value2 = temperatures[0]-273.15;
                this.Range["B" + (number + 7)].Value2 = cmeSim.getSaturationPressure();
                this.Range["C" + (number + 7)].Value2 = 1.0;
                this.Range["D" + (number + 7)].Value2 = "saturation point";
                this.Range["E" + (number + 7)].Value2 = cmeSim.getZsaturation();
                this.Range["G" + (number + 7)].Value2 = cmeSim.getSaturationIsoThermalCompressibility();

            }
            else if (PVTcalcCombobox.SelectedItem == "CVD")
            {
                calc1RowTitle.Value2 = "relativeVolume";
                calc2RowTitle.Value2 = "Liquid volume (% of Vsat)";
                calc3RowTitle.Value2 = "Z-gas";
                calc4RowTitle.Value2 = "Z-mix";
                calc5RowTitle.Value2 = "cummulative mole% depleted";

                ConstantVolumeDepletion cmdSim = new ConstantVolumeDepletion(thermoSystem);
                cmdSim.setTemperature(temperatures[0]);
                cmdSim.setPressures(pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (double val in cmdSim.getRelativeVolume())
                {
                    this.Range["C" + (number + 7)].Value2 = val;
                    this.Range["D" + (number + 7)].Value2 = cmdSim.getLiquidRelativeVolume()[number];
                    this.Range["E" + (number + 7)].Value2 = cmdSim.getZgas()[number];
                    this.Range["F" + (number + 7)].Value2 = cmdSim.getZmix()[number];
                    this.Range["G" + (number + 7)].Value2 = cmdSim.getCummulativeMolePercDepleted()[number];
                    number++;
                }
                this.Range["A" + (number + 7)].Value2 = temperatures[0]-273.15;
                this.Range["B" + (number + 7)].Value2 = cmdSim.getSaturationPressure();
                this.Range["C" + (number + 7)].Value2 = 1.0;
                this.Range["E" + (number + 7)].Value2 = cmdSim.getZsaturation();
                this.Range["F" + (number + 7)].Value2 = cmdSim.getZsaturation();
                this.Range["G" + (number + 7)].Value2 = "saturation point";
            }

            else if (PVTcalcCombobox.SelectedItem == "Differential liberation")
            {
                calc1RowTitle.Value2 = "Bo";
                    calc2RowTitle.Value2 = "Bg";
                     calc3RowTitle.Value2 = "Rs";
                      calc4RowTitle.Value2 = "gas gravity";
                       calc5RowTitle.Value2 = "Zgas";
                     calc6RowTitle.Value2 = "gas standard volume";

                DifferentialLiberation difLibSIm = new DifferentialLiberation(thermoSystem);
                difLibSIm.setTemperature(temperatures[0]);
                difLibSIm.setPressures(pressures);
                difLibSIm.runCalc();
                number = 0;
                foreach (double val in difLibSIm.getBo())
                {
                    this.Range["C" + (number + 7)].Value2 = val;
                    this.Range["D" + (number + 7)].Value2 = difLibSIm.getBg()[(number)];
                    this.Range["E" + (number + 7)].Value2 = difLibSIm.getRs()[(number)];
                    this.Range["F" + (number + 7)].Value2 = difLibSIm.getRelGasGravity()[(number)];
                    this.Range["G" + (number + 7)].Value2 = difLibSIm.getZgas()[(number)];
                    this.Range["H" + (number + 7)].Value2 = difLibSIm.getGasStandardVolume()[(number)];
                    number++;
                }
            }
            else if (PVTcalcCombobox.SelectedItem == "Wax content")
            {
                calc1RowTitle.Value2 = "wt% wax";
                calc2RowTitle.Value2 = "";
                calc3RowTitle.Value2 = "";
                calc4RowTitle.Value2 = "";
                calc5RowTitle.Value2 = "";

                WaxFractionSim cmdSim = new WaxFractionSim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (double val in cmdSim.getWaxFraction())
                {
                    this.Range["C" + (number + 7)].Value2 = val*100.0;
                    number++;
                }
            }
            else if (PVTcalcCombobox.SelectedItem == "Viscosity")
            {
                calc1RowTitle.Value2 = "Gas viscosity [kg/m*sec]";
                calc2RowTitle.Value2 = "Oil viscosity [kg/m*sec]";
                calc3RowTitle.Value2 = "Aqueous viscosity [kg/m*sec]";
                calc4RowTitle.Value2 = "";
                calc5RowTitle.Value2 = "";

                ViscositySim cmdSim = new ViscositySim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (double val in cmdSim.getGasViscosity())
                {
                    this.Range["C" + (number + 7)].Value2 = val;
                    this.Range["D" + (number + 7)].Value2 = cmdSim.getOilViscosity()[number];
                    this.Range["E" + (number + 7)].Value2 = cmdSim.getAqueousViscosity()[number];
                    number++;
                }
            }
            else if (PVTcalcCombobox.SelectedItem == "Viscosity oil/wax")
            {
                calc1RowTitle.Value2 = "share rate [1/sec]";
                calc2RowTitle.Value2 = "wt% wax";
                calc3RowTitle.Value2 = "Oil viscosity [kg/m*sec]";
                calc4RowTitle.Value2 = "Oil/wax viscosity [kg/m*sec]";
                calc5RowTitle.Value2 = "";
                calc6RowTitle.Value2 = "";


                double[] shareRate = new double[number];

                number = 0;
                foreach (Excel.Range r in column3range.Cells)
                {
                    string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    shareRate[number] = r.Value2;
                        number++;
                    }
                }



                ViscosityWaxOilSim cmdSim = new ViscosityWaxOilSim(thermoSystem);
                cmdSim.setTemperaturesAndPressures(temperatures, pressures);
                cmdSim.setShareRate(shareRate);
                cmdSim.runCalc();
                number = 0;
                foreach (double val in cmdSim.getGasViscosity())
                {
                    this.Range["D" + (number + 7)].Value2 = cmdSim.getWaxFraction()[number]*100;
                    this.Range["E" + (number + 7)].Value2 = cmdSim.getOilViscosity()[number];
                    this.Range["F" + (number + 7)].Value2 = cmdSim.getOilwaxDispersionViscosity()[number];
                    number++;
                }
            }
            else if (PVTcalcCombobox.SelectedItem == "Swelling test")
            {
                calc2RowTitle.Value2 = "Relative volume";
                int a2 = Convert.ToInt32((String)selectFluidCombobox.SelectedItem.ToString());
                SystemInterface gasInjectionThermoSystem = NeqSimThermoSystem.getThermoSystem().readObject(a2);

                SwellingTest cmdSim = new SwellingTest((SystemInterface) thermoSystem.clone());
                cmdSim.setInjectionGas(gasInjectionThermoSystem);
                cmdSim.setTemperature(temperatures[0]);
                cmdSim.setCummulativeMolePercentGasInjected(pressures);
                cmdSim.runCalc();
                number = 0;
                foreach (double val in cmdSim.getPressures())
                {
                    this.Range["C" + (number + 7)].Value2 = cmdSim.getPressures()[number];
                    this.Range["D" + (number + 7)].Value2 = cmdSim.getRelativeOilVolume()[number];
                    number++;
                }

            }
            rangeClear.Columns.AutoFit();
        }

        private void PVTcalcCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Range["B6"].Value2 = "Pressure [bara]";
            selectFluidCombobox.Visible = false;
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
            else if (PVTcalcCombobox.SelectedItem == "Viscosity oil/wax")
            {
                calc1RowTitle.Value2 = "Share rate [1/sec]";
            }
            else if (PVTcalcCombobox.SelectedItem == "Swelling test")
            {
                this.Range["B6"].Value2 = "GasInjected (fraction of oil)";
                this.Range["C6"].Value2 = "Pressure [bara]";
                selectFluidCombobox.Visible = true;
                this.Range["A4"].Value2 = "Injection gas";
                calc2RowTitle.Value2 = "Relative volume";
            }
            else if (PVTcalcCombobox.SelectedItem == "Slim tube")
            {
                selectFluidCombobox.Visible = true;
                this.Range["A4"].Value2 = "Injection gas";
            }

            rangeClear.Columns.AutoFit();

        }

    }
}
