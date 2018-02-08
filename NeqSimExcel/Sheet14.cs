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
    public partial class Sheet14
    {
        Excel.Range statusRange;
        String oldUnit = "mmol/kgWater";

        private void Sheet14_Startup(object sender, System.EventArgs e)
        {
            statusRange = this.Range["F20"];
        }

        private void Sheet14_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.unitComboBox.SelectedIndexChanged += new System.EventHandler(this.unitComboBox_SelectedIndexChanged);
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.calcpHButton.Click += new System.EventHandler(this.calcpHButton_Click);
            this.initWaterButton.Click += new System.EventHandler(this.initWaterButton_Click);
            this.Startup += new System.EventHandler(this.Sheet14_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet14_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            initWaterButton_Click(sender, e);
            statusRange.Value2 = "converting to electrolyte model....please wait...";
            SystemInterface thermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem();
            thermoSystem.setTotalFlowRate(this.Range["K12"].Value2, "MSm^3/day");
            thermoSystem = thermoSystem.setModel("Electrolyte-CPA-EOS-statoil");
            thermoSystem.setTemperature(273.15+this.Range["F9"].Value2);
            thermoSystem.setPressure(this.Range["F8"].Value2);
            statusRange.Value2 = "adding ions....please wait...";
            if (unitComboBox.SelectedItem.Equals("mg/kgWater"))
            {
                convertmgtommol();
            }
            double factor = this.Range["K13"].Value2*1000.0/(60*60*24.0);//kg/m^3*sec/day
            thermoSystem.addComponent("water", 1.0 * factor, "kg/sec");
            if (this.Range["B2"].Value2 > 0) thermoSystem.addComponent("Na+", this.Range["B2"].Value2 / 1000.0 * factor);
            if (this.Range["B3"].Value2 > 0) thermoSystem.addComponent("K+", this.Range["B3"].Value2 / 1000.0 * factor);
            if (this.Range["B4"].Value2 > 0) thermoSystem.addComponent("Mg++", this.Range["B4"].Value2 / 1000.0 * factor);
            if (this.Range["B5"].Value2 > 0) thermoSystem.addComponent("Ca++", this.Range["B5"].Value2 / 1000.0 * factor);
            if (this.Range["B6"].Value2 > 0) thermoSystem.addComponent("Ba++", this.Range["B6"].Value2 / 1000.0 * factor);
            if (this.Range["B7"].Value2 > 0) thermoSystem.addComponent("Sr++", this.Range["B7"].Value2 / 1000.0 * factor);
            if (this.Range["B8"].Value2 > 0) thermoSystem.addComponent("Fe++", this.Range["B8"].Value2 / 1000.0 * factor);
            if (this.Range["B9"].Value2 > 0) thermoSystem.addComponent("Cl-", this.Range["B9"].Value2 / 1000.0 * factor);
            if (this.Range["B10"].Value2 > 0) thermoSystem.addComponent("Br-", this.Range["B10"].Value2 / 1000.0 * factor);
            if (this.Range["B11"].Value2 > 0) thermoSystem.addComponent("SO4--", this.Range["B11"].Value2 / 1000.0 * factor);

            if (this.Range["B13"].Value2 > 0) thermoSystem.addComponent("OH-", this.Range["B13"].Value2 / 1000.0*factor);

            if (unitComboBox.SelectedItem.Equals("mg/kgWater"))
            {
                convertmmoltomg();
            }
            statusRange.Value2 = "initializing chemical reactions....please wait...";
            thermoSystem.chemicalReactionInit();
            thermoSystem.createDatabase(true);
            thermoSystem.setMixingRule(10);
            thermoSystem.setMultiPhaseCheck(true);
            
            ThermodynamicOperations testOps = new ThermodynamicOperations(thermoSystem);
            try
            {
                statusRange.Value2 = "TPflash running....please wait...";
                testOps.TPflash();
                //thermoSystem.display();
                if (thermoSystem.getPhase(0).hasComponent("CO2"))
                {
                    this.Range["F13"].Value2 = thermoSystem.getPhase(0).getComponent("CO2").getx() * 100;
                }
                else
                {

                    this.Range["F13"].Value2 = 0;
                }
                if (thermoSystem.getPhase(0).hasComponent("H2S"))  {
                    this.Range["F14"].Value2 = thermoSystem.getPhase(0).getComponent("H2S").getx() * 100;
                      }
                else
                {

                    this.Range["F14"].Value2 = 0;
                }
                this.Range["E22"].Value2 = "pH";
                this.Range["F22"].Value2 = thermoSystem.getPhase("aqueous").getpH();
                if (thermoSystem.getPhase(0).hasComponent("CO2")) this.Range["F13"].Value2 = thermoSystem.getPhase(0).getComponent("CO2").getx() * 100;
                if (thermoSystem.getPhase(0).hasComponent("H2S")) this.Range["F14"].Value2 = thermoSystem.getPhase(0).getComponent("H2S").getx() * 100;

                this.Range["E22"].Value2 = "pH";
                this.Range["F22"].Value2 = thermoSystem.getPhase("aqueous").getpH();
                //thermoSystem.display();
                statusRange.Value2 = "calculationg ion composition....please wait...";
                testOps.calcIonComposition(thermoSystem.getPhaseNumberOfPhase("aqueous"));
                string[][] ionComp = testOps.getResultTable();
                int numb = ionComp.Length;
                for (int i = 0; i < ionComp.Length; i++)
                {
                    string name1 = "E" + (i + 25).ToString();
                    this.Range[name1].Value2 = ionComp[i][0];

                    string name2 = "F" + (i + 25).ToString();
                    this.Range[name2].Value2 = ionComp[i][1];

                    string name3 = "G" + (i + 25).ToString();
                    this.Range[name3].Value2 = ionComp[i][2];

                    string name4 = "H" + (i + 25).ToString();
                    this.Range[name4].Value2 = ionComp[i][3];

                }

                int startNumb = 30 + numb;
                statusRange.Value2 = "calculationg scale potential....please wait...";

                testOps.checkScalePotential(thermoSystem.getPhaseNumberOfPhase("aqueous"));
                string[][] ionComp2 = testOps.getResultTable();
                for (int i = 0; i < ionComp2.Length; i++)
                {
                    string name1 = "E" + (i + startNumb).ToString();
                    this.Range[name1].Value2 = ionComp2[i][0];

                    string name2 = "F" + (i + startNumb).ToString();
                    this.Range[name2].Value2 = ionComp2[i][1];


                }

                int writeStartCell = 25, writeEndCell = 33;

                var table = thermoSystem.createTable("fluid");
                int rows = table.Length;
                int columns = table[1].Length;
                writeEndCell = writeStartCell + rows;

                var startCell = Cells[writeStartCell, 10];
                var endCell = Cells[writeEndCell - 1, columns + 9];
                var writeRange = this.Range[startCell, endCell];

                writeStartCell += rows + 3;
                //writeRange.Value2 = table;

                var data = new object[rows, columns];
                for (var row = 1; row <= rows; row++)
                {
                    for (var column = 1; column <= columns; column++)
                    {
                        data[row - 1, column - 1] = table[row - 1][column - 1];
                    }
                }

                writeRange.Value2 = data;

                if (saveFluidCheckBox.Checked)
                {
                    NeqSimThermoSystem.setThermoSystem(thermoSystem);
                }

                statusRange.Value2 = "finished calculations.";

            }
            catch (Exception ex)
            {
                statusRange.Value2 = "error occured during calculations..";
                //   ex.Message();
                return;
            }

        }

        private void calcpHButton_Click(object sender, EventArgs e)
        {
            Excel.Range rangeClear = this.Range["E22", "L100"];
            rangeClear.Clear();

            initWaterButton_Click(sender,e);
            if (unitComboBox.SelectedItem.Equals("mg/kgWater")){
                convertmgtommol();
            }
            statusRange.Value2 = "creating fluid....please wait...";

            SystemInterface thermoSystem = (SystemInterface)new SystemElectrolyteCPAstatoil(273.15+this.Range["F9"].Value2, this.Range["F8"].Value2);
            //SystemInterface thermoSystem = (SystemInterface)new SystemSrkCPAstatoil(298, 10);
            if (this.Range["F12"].Value2 > 0)  thermoSystem.addComponent("methane", 1.0 * this.Range["F12"].Value2 / 1.0);
            if (this.Range["F13"].Value2 > 0)
            {
                thermoSystem.addComponent("CO2", 1.0 * this.Range["F13"].Value2 / 1.0 + this.Range["B13"].Value2 / 1000.0);
            }
            if (this.Range["F14"].Value2 > 0)
            {
                thermoSystem.addComponent("H2S", 1.0 * this.Range["F14"].Value2 / 1.0);
            }

            thermoSystem.addComponent("water", 1.0, "kg/sec");
            if (this.Range["B2"].Value2 > 0) thermoSystem.addComponent("Na+", this.Range["B2"].Value2/1000.0);
            if (this.Range["B3"].Value2 > 0) thermoSystem.addComponent("K+", this.Range["B3"].Value2 / 1000.0);
            if (this.Range["B4"].Value2 > 0) thermoSystem.addComponent("Mg++", this.Range["B4"].Value2 / 1000.0);
            if (this.Range["B5"].Value2 > 0) thermoSystem.addComponent("Ca++", this.Range["B5"].Value2 / 1000.0);
            if (this.Range["B6"].Value2 > 0) thermoSystem.addComponent("Ba++", this.Range["B6"].Value2 / 1000.0);
            if (this.Range["B7"].Value2 > 0) thermoSystem.addComponent("Sr++", this.Range["B7"].Value2 / 1000.0);
            if (this.Range["B8"].Value2 > 0) thermoSystem.addComponent("Fe++", this.Range["B8"].Value2 / 1000.0);
            if (this.Range["B9"].Value2 > 0) thermoSystem.addComponent("Cl-", this.Range["B9"].Value2 / 1000.0);
            if (this.Range["B10"].Value2 > 0) thermoSystem.addComponent("Br-", this.Range["B10"].Value2 / 1000.0);
            if (this.Range["B11"].Value2 > 0) thermoSystem.addComponent("SO4--", this.Range["B11"].Value2 / 1000.0);

            if (this.Range["B13"].Value2 > 0) thermoSystem.addComponent("OH-", this.Range["B13"].Value2 / 1000.0);
            
            if (unitComboBox.SelectedItem.Equals("mg/kgWater")){
                convertmmoltomg();
            }
            statusRange.Value2 = "initializing chemical reactions....please wait...";
            thermoSystem.chemicalReactionInit();
            thermoSystem.createDatabase(true);
            thermoSystem.setMixingRule(10);
           // thermoSystem.setMultiPhaseCheck(true);

            ThermodynamicOperations testOps = new ThermodynamicOperations(thermoSystem);
             try
             {
                 statusRange.Value2 = "TPflash running....please wait...";
                 testOps.TPflash();
                 if(thermoSystem.getPhase(0).hasComponent("CO2")) this.Range["F13"].Value2 = thermoSystem.getPhase(0).getComponent("CO2").getx()*100;
                 if (thermoSystem.getPhase(0).hasComponent("H2S")) this.Range["F14"].Value2 = thermoSystem.getPhase(0).getComponent("H2S").getx() * 100;
                 
                 this.Range["E22"].Value2 = "pH";
                 this.Range["F22"].Value2 = thermoSystem.getPhase("aqueous").getpH();
                 //thermoSystem.display();
                 statusRange.Value2 = "calculationg ion composition....please wait...";
                 testOps.calcIonComposition(thermoSystem.getPhaseNumberOfPhase("aqueous"));
                 string[][] ionComp = testOps.getResultTable();
                 int numb = ionComp.Length;
                 for (int i = 0; i < ionComp.Length; i++)
                 {
                     string name1 = "E" + (i + 25).ToString();
                     this.Range[name1].Value2 =  ionComp[i][0];

                     string name2 = "F" + (i + 25).ToString();
                     this.Range[name2].Value2 = ionComp[i][1];

                     string name3 = "G" + (i + 25).ToString();
                     this.Range[name3].Value2 = ionComp[i][2];

                     string name4 = "H" + (i + 25).ToString();
                     this.Range[name4].Value2 = ionComp[i][3];



                 }
                 int startNumb = 30 + numb;
                 statusRange.Value2 = "calculationg scale potential....please wait...";
                 
                 testOps.checkScalePotential(thermoSystem.getPhaseNumberOfPhase("aqueous"));
                 string[][] ionComp2 = testOps.getResultTable();
                for (int i = 0; i < ionComp2.Length; i++)
                 {
                     string name1 = "E" + (i + startNumb).ToString();
                     this.Range[name1].Value2 = ionComp2[i][0];

                     string name2 = "F" + (i + startNumb).ToString();
                     this.Range[name2].Value2 = ionComp2[i][1];

                    
                 }
                statusRange.Value2 = "finished calculations.";
                // testOps.display();
             } 
             catch (Exception ex) {
                 statusRange.Value2 = "error calculating flash...";
         //   ex.Message();
        }
                
        }

        private void initWaterButton_Click(object sender, EventArgs e)
        {
            statusRange.Value2 = "initializing charge";
            if (unitComboBox.SelectedItem.Equals("mg/kgWater"))
            {
                convertmgtommol();
            }
            double totalCHarge = this.Range["B15"].Value2;
            if (totalCHarge < 0)
            {
                this.Range["B2"].Value2 = this.Range["B2"].Value2 - totalCHarge;
               // thermoSystem.addComponent("Na+", -totalCHarge / 1000.0);
            }
            if (totalCHarge > 0)
            {
                this.Range["B9"].Value2 = this.Range["B9"].Value2 + totalCHarge;
              //  thermoSystem.addComponent("Cl-", totalCHarge / 1000.0);
            }
            if (unitComboBox.SelectedItem.Equals("mg/kgWater"))
            {
                convertmmoltomg();
            }
        }

        private void convertmmoltomg()
        {
            this.Range["B2"].Value2 = this.Range["B2"].Value2 * 22.99; //Na+
            this.Range["B3"].Value2 = this.Range["B3"].Value2 * 39.1;
            this.Range["B4"].Value2 = this.Range["B4"].Value2 * 24.31;
            this.Range["B5"].Value2 = this.Range["B5"].Value2 * 40.08;
            this.Range["B6"].Value2 = this.Range["B6"].Value2 * 137.3;
            this.Range["B7"].Value2 = this.Range["B7"].Value2 * 87.62;
            this.Range["B8"].Value2 = this.Range["B8"].Value2 * 55.85;
            this.Range["B9"].Value2 = this.Range["B9"].Value2 * 35.45;
            this.Range["B10"].Value2 = this.Range["B10"].Value2 * 79.9;
            this.Range["B11"].Value2 = this.Range["B11"].Value2 * 96.07;
            this.Range["B13"].Value2 = this.Range["B13"].Value2 * 17.001;
        }

        private void convertmgtommol()
        {
            this.Range["B2"].Value2 = this.Range["B2"].Value2 / 22.99; //Na+
            this.Range["B3"].Value2 = this.Range["B3"].Value2 / 39.1;
            this.Range["B4"].Value2 = this.Range["B4"].Value2 / 24.31;
            this.Range["B5"].Value2 = this.Range["B5"].Value2 / 40.08;
            this.Range["B6"].Value2 = this.Range["B6"].Value2 / 137.3;
            this.Range["B7"].Value2 = this.Range["B7"].Value2 / 87.62;
            this.Range["B8"].Value2 = this.Range["B8"].Value2 / 55.85;
            this.Range["B9"].Value2 = this.Range["B9"].Value2 / 35.45;
            this.Range["B10"].Value2 = this.Range["B10"].Value2 / 79.9;
            this.Range["B11"].Value2 = this.Range["B11"].Value2 / 96.07;
            this.Range["B13"].Value2 = this.Range["B13"].Value2 / 17.001;
        }

        private void unitComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            this.Range["B1"].Value2 = unitComboBox.SelectedItem;

            if(unitComboBox.SelectedItem.Equals(oldUnit)) return;

            if (oldUnit.Equals("mmol/kgWater") && unitComboBox.SelectedItem.Equals("mg/kgWater"))
            {
                convertmmoltomg();
        

            }
            if (oldUnit.Equals("mg/kgWater") && unitComboBox.SelectedItem.Equals("mmol/kgWater"))
            {
                convertmgtommol();
            }
            oldUnit = unitComboBox.SelectedItem.ToString();

        }

    }
}
