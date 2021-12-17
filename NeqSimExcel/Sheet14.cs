using System;
using Microsoft.Office.Interop.Excel;
using neqsim.thermo.system;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet14
    {
        private string oldUnit = "mmol/kgWater";
        private Range statusRange;

        private void Sheet14_Startup(object sender, EventArgs e)
        {
            statusRange = Range["F20"];
            unitComboBox.SelectedIndex = 0;
        }

        private void Sheet14_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
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
            var rangeClear = Range["E22", "Y200"];
            rangeClear.Clear();
            initWaterButton_Click(sender, e);
            statusRange.Value2 = "converting to electrolyte model....please wait...";
            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            thermoSystem.setTotalFlowRate(Range["K12"].Value2, "MSm^3/day");
            thermoSystem = thermoSystem.setModel("Electrolyte-CPA-EOS-statoil");
            thermoSystem.setTemperature(273.15 + Range["F9"].Value2);
            thermoSystem.setPressure(Range["F8"].Value2);
            statusRange.Value2 = "adding ions....please wait...";
            if (unitComboBox.SelectedItem.Equals("mg/kgWater")) convertmgtommol();
            double factor = Range["K13"].Value2 * 1000.0 / (60 * 60 * 24.0); //kg/m^3*sec/day
            thermoSystem.addComponent("water", 1.0 * factor, "kg/sec");
            if (Range["B2"].Value2 > 0) thermoSystem.addComponent("Na+", Range["B2"].Value2 / 1000.0 * factor);
            if (Range["B3"].Value2 > 0) thermoSystem.addComponent("K+", Range["B3"].Value2 / 1000.0 * factor);
            if (Range["B4"].Value2 > 0) thermoSystem.addComponent("Mg++", Range["B4"].Value2 / 1000.0 * factor);
            if (Range["B5"].Value2 > 0) thermoSystem.addComponent("Ca++", Range["B5"].Value2 / 1000.0 * factor);
            if (Range["B6"].Value2 > 0) thermoSystem.addComponent("Ba++", Range["B6"].Value2 / 1000.0 * factor);
            if (Range["B7"].Value2 > 0) thermoSystem.addComponent("Sr++", Range["B7"].Value2 / 1000.0 * factor);
            if (Range["B8"].Value2 > 0) thermoSystem.addComponent("Fe++", Range["B8"].Value2 / 1000.0 * factor);
            if (Range["B9"].Value2 > 0) thermoSystem.addComponent("Cl-", Range["B9"].Value2 / 1000.0 * factor);
            if (Range["B10"].Value2 > 0) thermoSystem.addComponent("Br-", Range["B10"].Value2 / 1000.0 * factor);
            if (Range["B11"].Value2 > 0) thermoSystem.addComponent("SO4--", Range["B11"].Value2 / 1000.0 * factor);

            if (Range["B13"].Value2 > 0) thermoSystem.addComponent("OH-", Range["B13"].Value2 / 1000.0 * factor);

            if (unitComboBox.SelectedItem.Equals("mg/kgWater")) convertmmoltomg();
            statusRange.Value2 = "initializing chemical reactions....please wait...";
            thermoSystem.chemicalReactionInit();
            thermoSystem.createDatabase(true);
            thermoSystem.setMixingRule(10);
            thermoSystem.setMultiPhaseCheck(true);

            var testOps = new ThermodynamicOperations(thermoSystem);
            try
            {
                statusRange.Value2 = "TPflash running....please wait...";
                testOps.TPflash();
                //thermoSystem.display();
                if (thermoSystem.getPhase(0).hasComponent("CO2"))
                    Range["F13"].Value2 = thermoSystem.getPhase(0).getComponent("CO2").getx() * 100;
                else
                    Range["F13"].Value2 = 0;
                if (thermoSystem.getPhase(0).hasComponent("H2S"))
                    Range["F14"].Value2 = thermoSystem.getPhase(0).getComponent("H2S").getx() * 100;
                else
                    Range["F14"].Value2 = 0;
                Range["E22"].Value2 = "pH";
                Range["F22"].Value2 = thermoSystem.getPhase("aqueous").getpH();
                if (thermoSystem.getPhase(0).hasComponent("CO2"))
                    Range["F13"].Value2 = thermoSystem.getPhase(0).getComponent("CO2").getx() * 100;
                if (thermoSystem.getPhase(0).hasComponent("H2S"))
                    Range["F14"].Value2 = thermoSystem.getPhase(0).getComponent("H2S").getx() * 100;

                Range["E22"].Value2 = "pH";
                Range["F22"].Value2 = thermoSystem.getPhase("aqueous").getpH();
                //thermoSystem.display();
                statusRange.Value2 = "calculationg ion composition....please wait...";
                testOps.calcIonComposition(thermoSystem.getPhaseNumberOfPhase("aqueous"));
                var ionComp = testOps.getResultTable();
                var numb = ionComp.Length;
                for (var i = 0; i < ionComp.Length; i++)
                {
                    var name1 = "E" + (i + 25);
                    Range[name1].Value2 = ionComp[i][0];

                    var name2 = "F" + (i + 25);
                    Range[name2].Value2 = ionComp[i][1];

                    var name3 = "G" + (i + 25);
                    Range[name3].Value2 = ionComp[i][2];

                    var name4 = "H" + (i + 25);
                    Range[name4].Value2 = ionComp[i][3];
                }

                var startNumb = 30 + numb;
                statusRange.Value2 = "calculationg scale potential....please wait...";

                testOps.checkScalePotential(thermoSystem.getPhaseNumberOfPhase("aqueous"));
                var ionComp2 = testOps.getResultTable();
                for (var i = 0; i < ionComp2.Length; i++)
                {
                    var name1 = "E" + (i + startNumb);
                    Range[name1].Value2 = ionComp2[i][0];

                    var name2 = "F" + (i + startNumb);
                    Range[name2].Value2 = ionComp2[i][1];
                }

                int writeStartCell = 25, writeEndCell = 33;

                var table = thermoSystem.createTable("fluid");
                var rows = table.Length;
                var columns = table[1].Length;
                writeEndCell = writeStartCell + rows;

                var startCell = Cells[writeStartCell, 10];
                var endCell = Cells[writeEndCell - 1, columns + 9];
                var writeRange = Range[startCell, endCell];

                writeStartCell += rows + 3;
                //writeRange.Value2 = table;

                var data = new object[rows, columns];
                for (var row = 1; row <= rows; row++)
                for (var column = 1; column <= columns; column++)
                    data[row - 1, column - 1] = table[row - 1][column - 1];

                writeRange.Value2 = data;

                if (saveFluidCheckBox.Checked) NeqSimThermoSystem.setThermoSystem(thermoSystem);

                statusRange.Value2 = "finished calculations.";
            }
            catch (Exception ex)
            {
                statusRange.Value2 = "error occured during calculations.."+ ex.Message;
                //   ex.Message();
            }
        }

        private void calcpHButton_Click(object sender, EventArgs e)
        {
            var rangeClear = Range["E22", "Y200"];
            rangeClear.Clear();

            initWaterButton_Click(sender, e);
            if (unitComboBox.SelectedItem.Equals("mg/kgWater")) convertmgtommol();
            statusRange.Value2 = "creating fluid....please wait...";

            var thermoSystem =
                (SystemInterface) new SystemElectrolyteCPAstatoil(273.15 + Range["F9"].Value2, Range["F8"].Value2);
            //SystemInterface thermoSystem = (SystemInterface)new SystemSrkCPAstatoil(298, 10);
            if (Range["F12"].Value2 > 0) thermoSystem.addComponent("methane", 1.0 * Range["F12"].Value2 / 1.0);
            if (Range["F13"].Value2 > 0)
                thermoSystem.addComponent("CO2", 1.0 * Range["F13"].Value2 / 1.0 + Range["B13"].Value2 / 1000.0);
            if (Range["F14"].Value2 > 0) thermoSystem.addComponent("H2S", 1.0 * Range["F14"].Value2 / 1.0);

            thermoSystem.addComponent("water", 1.0, "kg/sec");
            if (Range["B2"].Value2 > 0) thermoSystem.addComponent("Na+", Range["B2"].Value2 / 1000.0);
            if (Range["B3"].Value2 > 0) thermoSystem.addComponent("K+", Range["B3"].Value2 / 1000.0);
            if (Range["B4"].Value2 > 0) thermoSystem.addComponent("Mg++", Range["B4"].Value2 / 1000.0);
            if (Range["B5"].Value2 > 0) thermoSystem.addComponent("Ca++", Range["B5"].Value2 / 1000.0);
            if (Range["B6"].Value2 > 0) thermoSystem.addComponent("Ba++", Range["B6"].Value2 / 1000.0);
            if (Range["B7"].Value2 > 0) thermoSystem.addComponent("Sr++", Range["B7"].Value2 / 1000.0);
            if (Range["B8"].Value2 > 0) thermoSystem.addComponent("Fe++", Range["B8"].Value2 / 1000.0);
            if (Range["B9"].Value2 > 0) thermoSystem.addComponent("Cl-", Range["B9"].Value2 / 1000.0);
            if (Range["B10"].Value2 > 0) thermoSystem.addComponent("Br-", Range["B10"].Value2 / 1000.0);
            if (Range["B11"].Value2 > 0) thermoSystem.addComponent("SO4--", Range["B11"].Value2 / 1000.0);

            if (Range["B13"].Value2 > 0) thermoSystem.addComponent("OH-", Range["B13"].Value2 / 1000.0);

            if (unitComboBox.SelectedItem.Equals("mg/kgWater")) convertmmoltomg();
            statusRange.Value2 = "initializing chemical reactions....please wait...";
            thermoSystem.chemicalReactionInit();
            thermoSystem.createDatabase(true);
            thermoSystem.setMixingRule(10);
            // thermoSystem.setMultiPhaseCheck(true);

            var testOps = new ThermodynamicOperations(thermoSystem);
            try
            {
                statusRange.Value2 = "TPflash running....please wait...";
                testOps.TPflash();
                if (thermoSystem.getPhase(0).hasComponent("CO2"))
                    Range["F13"].Value2 = thermoSystem.getPhase(0).getComponent("CO2").getx() * 100;
                if (thermoSystem.getPhase(0).hasComponent("H2S"))
                    Range["F14"].Value2 = thermoSystem.getPhase(0).getComponent("H2S").getx() * 100;

                Range["E22"].Value2 = "pH";
                Range["F22"].Value2 = thermoSystem.getPhase("aqueous").getpH();
                //thermoSystem.display();
                statusRange.Value2 = "calculationg ion composition....please wait...";
                testOps.calcIonComposition(thermoSystem.getPhaseNumberOfPhase("aqueous"));
                var ionComp = testOps.getResultTable();
                var numb = ionComp.Length;
                for (var i = 0; i < ionComp.Length; i++)
                {
                    var name1 = "E" + (i + 25);
                    Range[name1].Value2 = ionComp[i][0];

                    var name2 = "F" + (i + 25);
                    Range[name2].Value2 = ionComp[i][1];

                    var name3 = "G" + (i + 25);
                    Range[name3].Value2 = ionComp[i][2];

                    var name4 = "H" + (i + 25);
                    Range[name4].Value2 = ionComp[i][3];
                }

                var startNumb = 30 + numb;
                statusRange.Value2 = "calculationg scale potential....please wait...";

                testOps.checkScalePotential(thermoSystem.getPhaseNumberOfPhase("aqueous"));
                var ionComp2 = testOps.getResultTable();
                for (var i = 0; i < ionComp2.Length; i++)
                {
                    var name1 = "E" + (i + startNumb);
                    Range[name1].Value2 = ionComp2[i][0];

                    var name2 = "F" + (i + startNumb);
                    Range[name2].Value2 = ionComp2[i][1];
                }

                statusRange.Value2 = "finished calculations.";
                // testOps.display();
            }
            catch (Exception ex)
            {
                statusRange.Value2 = "error calculating flash..." + ex.Message;
            }
        }

        private void initWaterButton_Click(object sender, EventArgs e)
        {
            statusRange.Value2 = "initializing charge";
            if (unitComboBox.SelectedItem.Equals("mg/kgWater")) convertmgtommol();
            if (unitComboBox.SelectedItem.Equals("mg/litre")) convertmglitretommol();
            double totalCHarge = Range["B15"].Value2;
            if (totalCHarge < 0) Range["B2"].Value2 = Range["B2"].Value2 - totalCHarge;
            if (totalCHarge > 0) Range["B9"].Value2 = Range["B9"].Value2 + totalCHarge;
            if (unitComboBox.SelectedItem.Equals("mg/kgWater")) convertmmoltomg();
        }

        private void convertmmoltomg()
        {
            Range["B2"].Value2 = Range["B2"].Value2 * 22.99; //Na+
            Range["B3"].Value2 = Range["B3"].Value2 * 39.1;
            Range["B4"].Value2 = Range["B4"].Value2 * 24.31;
            Range["B5"].Value2 = Range["B5"].Value2 * 40.08;
            Range["B6"].Value2 = Range["B6"].Value2 * 137.3;
            Range["B7"].Value2 = Range["B7"].Value2 * 87.62;
            Range["B8"].Value2 = Range["B8"].Value2 * 55.85;
            Range["B9"].Value2 = Range["B9"].Value2 * 35.45;
            Range["B10"].Value2 = Range["B10"].Value2 * 79.9;
            Range["B11"].Value2 = Range["B11"].Value2 * 96.07;
            Range["B13"].Value2 = Range["B13"].Value2 * 17.001;
        }
        
             private void convertmglitretommol()
        {
            //this method needs changes further work here

            //double densitySolution = 1100.0;//needs calculation
            
            //double kgwaterperkgsolution = densitySolution-

            Range["B2"].Value2 = Range["B2"].Value2 / 22.99;
            Range["B3"].Value2 = Range["B3"].Value2 / 39.1;
            Range["B4"].Value2 = Range["B4"].Value2 / 24.31;
            Range["B5"].Value2 = Range["B5"].Value2 / 40.08;
            Range["B6"].Value2 = Range["B6"].Value2 / 137.3;
            Range["B7"].Value2 = Range["B7"].Value2 / 87.62;
            Range["B8"].Value2 = Range["B8"].Value2 / 55.85;
            Range["B9"].Value2 = Range["B9"].Value2 / 35.45;
            Range["B10"].Value2 = Range["B10"].Value2 / 79.9;
            Range["B11"].Value2 = Range["B11"].Value2 / 96.07;
            Range["B13"].Value2 = Range["B13"].Value2 / 17.001;
        }

        private void convertmgtommol()
        {
            Range["B2"].Value2 = Range["B2"].Value2 / 22.99; //Na+
            Range["B3"].Value2 = Range["B3"].Value2 / 39.1;
            Range["B4"].Value2 = Range["B4"].Value2 / 24.31;
            Range["B5"].Value2 = Range["B5"].Value2 / 40.08;
            Range["B6"].Value2 = Range["B6"].Value2 / 137.3;
            Range["B7"].Value2 = Range["B7"].Value2 / 87.62;
            Range["B8"].Value2 = Range["B8"].Value2 / 55.85;
            Range["B9"].Value2 = Range["B9"].Value2 / 35.45;
            Range["B10"].Value2 = Range["B10"].Value2 / 79.9;
            Range["B11"].Value2 = Range["B11"].Value2 / 96.07;
            Range["B13"].Value2 = Range["B13"].Value2 / 17.001;
        }

        private void unitComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Range["B1"].Value2 = unitComboBox.SelectedItem;

            if (unitComboBox.SelectedItem.Equals(oldUnit)) return;

            if (oldUnit.Equals("mmol/kgWater") && unitComboBox.SelectedItem.Equals("mg/kgWater")) convertmmoltomg();
            if (oldUnit.Equals("mg/kgWater") && unitComboBox.SelectedItem.Equals("mmol/kgWater")) convertmgtommol();
            oldUnit = unitComboBox.SelectedItem.ToString();
        }
    }
}