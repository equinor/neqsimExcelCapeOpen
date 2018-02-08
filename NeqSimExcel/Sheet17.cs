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
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Collections;
using System.Security.Principal;
using MySql.Data;

namespace NeqSimExcel
{
    public partial class Sheet17
    {
        SystemInterface gasThermoSystem, liqThermoSystem;
        Excel.Range statusRange = null;

        private void Sheet17_Startup(object sender, System.EventArgs e)
        {
            statusRange = this.Range["D14"];

            try
            {
                DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter test = new DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter();
               // NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter test = new NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter test = new NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfo1TableAdapter test = new neqsimdatabaseDataSetTableAdapters.fluidinfo1TableAdapter();

                string userName = WindowsIdentity.GetCurrent().Name;
                userName = userName.Replace("STATOIL-NET\\", "");
                userName = userName.Replace("WIN-NTNU-NO\\", "");
                userName = userName.ToLower();


                DatabaseConnection.NeqSimDatabaseSet.fluidinfoDataTable tt = test.GetDataBy(userName);
                List<string> names = new List<string>();
                //names.Add("CPApackage");
                //names.Add(WindowsIdentity.GetCurrent().Name);
                foreach (DatabaseConnection.NeqSimDatabaseSet.fluidinfoRow row in tt.Rows)
                {
                    names.Add(row.ID.ToString());
                    fluidListNameComboBoxGas.Items.Add(row.ID.ToString());
                    fluidListNameComboBoxLiquid.Items.Add(row.ID.ToString());
                }
                //   packageNames = names.ToArray();
                //   fluidListNameComboBox.Items.Add(names.ToList());
                fluidListNameComboBoxGas.SelectedIndex = 0;
                fluidListNameComboBoxLiquid.SelectedIndex = 0;

                int gasNumb = Convert.ToInt32((String)fluidListNameComboBoxGas.SelectedItem.ToString());
                gasThermoSystem = NeqSimThermoSystem.getThermoSystem();
                gasThermoSystem = gasThermoSystem.readObject(gasNumb);

                int liqNumb = Convert.ToInt32((String)fluidListNameComboBoxLiquid.SelectedItem.ToString());
                liqThermoSystem = NeqSimThermoSystem.getThermoSystem();
                liqThermoSystem = liqThermoSystem.readObject(liqNumb);
            

            }
            catch (Exception excet)
            {
                Console.WriteLine("Error " + excet.Message);
            }
        }

        private void Sheet17_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.fluidListNameComboBoxGas.SelectedIndexChanged += new System.EventHandler(this.fluidListNameComboBoxGas_SelectedIndexChanged);
            this.calculateButton.Click += new System.EventHandler(this.calculateButton_Click);
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            this.linkLabel1.Click += new System.EventHandler(this.linkLabel1_Click);
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            this.Startup += new System.EventHandler(this.Sheet17_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet17_Shutdown);

        }

        #endregion

        private void fluidListNameComboBoxGas_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void calculateButton_Click(object sender, EventArgs e)
        {
            statusRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            this.Range["F2", "H100"].Clear();

            statusRange.Value2 = "reading fluids...";
            int gasNumb = Convert.ToInt32((String)fluidListNameComboBoxGas.SelectedItem.ToString());
            gasThermoSystem = NeqSimThermoSystem.getThermoSystem();
            gasThermoSystem = gasThermoSystem.readObject(gasNumb);
            gasThermoSystem.setTemperature(this.Range["B6"].Value2 + 273.15);
            gasThermoSystem.setPressure(this.Range["B7"].Value2);
            ThermodynamicOperations gasOps = new ThermodynamicOperations(gasThermoSystem);
            gasOps.TPflash();
            if(!gasThermoSystem.hasPhaseType("gas")){
                statusRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                statusRange.Value2 = "no gas at given temperature and pressure...check gas";
                return;
            }

            statusRange.Value2 = "reading gas...ok";

            int liqNumb = Convert.ToInt32((String)fluidListNameComboBoxLiquid.SelectedItem.ToString());
            liqThermoSystem = NeqSimThermoSystem.getThermoSystem();
            liqThermoSystem = liqThermoSystem.readObject(liqNumb);
            liqThermoSystem.setTotalFlowRate(this.Range["B11"].Value2, "kg/sec");
            liqThermoSystem.init(0);
            liqThermoSystem.init(1);
            liqThermoSystem.setMultiPhaseCheck(true);
            liqThermoSystem.setTemperature(this.Range["B6"].Value2 + 273.15);
            liqThermoSystem.setPressure(this.Range["B7"].Value2);
            ThermodynamicOperations liqOps = new ThermodynamicOperations(liqThermoSystem);
            liqOps.TPflash();
            if(!(liqThermoSystem.hasPhaseType("aqueous") || liqThermoSystem.hasPhaseType("liquid"))){
                statusRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                statusRange.Value2 = "no liquid phase at given temperature and pressure...check liquid";
                return;
            }

  
            statusRange.Value2 = "reading liquid...ok";
            if (!liqThermoSystem.getPhase(0).hasComponent("water"))
            {
                statusRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                statusRange.Value2 = "liquid phase must contain water... stopping calcuations";
                return;
            }
            ThermodynamicOperations testOps = new ThermodynamicOperations(liqThermoSystem);


            double oldTime = 0.0;
            Excel.Range range = this.Range["E2", "E100"];
            int number = 0;

            statusRange.Value2 = "calculations running.....";
            foreach (Excel.Range r in range.Cells)
            {
                string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    double time = this.Range[("E" + (number + 2).ToString())].Value2;
                    gasThermoSystem.setTotalFlowRate(this.Range["B5"].Value2 * (time - oldTime) / 24.0 * 3600, "MSm3/day");
                    oldTime = time;
                    gasThermoSystem.init(1);
                    liqThermoSystem.addFluid(gasThermoSystem);
                    testOps.TPflash();
                    int phaseN = liqThermoSystem.getNumberOfPhases();
                    if (phaseN > 1)
                    {
                        this.Range["F" + (number + 2).ToString()].Value2 = liqThermoSystem.getPhase(phaseN - 1).getWtFrac(liqThermoSystem.getPhase(phaseN - 1).getComponent("water").getComponentNumber())*100.0;
                        this.Range["H" + (number + 2).ToString()].Value2 = liqThermoSystem.getPhase(0).getComponent("water").getx() * 1e6;
                        this.Range["G" + (number + 2).ToString()].Value2 = liqThermoSystem.getPhase(phaseN - 1).getNumberOfMolesInPhase() * liqThermoSystem.getPhase(phaseN - 1).getMolarMass();
                        liqThermoSystem = liqThermoSystem.phaseToSystem(phaseN - 1);
                    }
                    testOps = new ThermodynamicOperations(liqThermoSystem);
                    // System.out.println("nuber of moles " + testSystem.getNumberOfMoles() + " moleFrac MEG " + testSystem.getPhase(0).getComponent("MEG").getx());
                    number++;
                }
            }
            statusRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
            statusRange.Value2 = "finished calcuatons";
        }


        private void linkLabel1_Click(object sender, EventArgs e)
        {
            int gasNumb = Convert.ToInt32((String)fluidListNameComboBoxGas.SelectedItem.ToString());
            gasThermoSystem = NeqSimThermoSystem.getThermoSystem();
            gasThermoSystem = gasThermoSystem.readObject(gasNumb);
            gasThermoSystem.setTemperature(this.Range["B6"].Value2 + 273.15);
            gasThermoSystem.setPressure(this.Range["B7"].Value2);
            ThermodynamicOperations testOps = new ThermodynamicOperations(gasThermoSystem);
            testOps.TPflash();
            gasThermoSystem.display();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            int liqNumb = Convert.ToInt32((String)fluidListNameComboBoxLiquid.SelectedItem.ToString());
            liqThermoSystem = NeqSimThermoSystem.getThermoSystem();
            liqThermoSystem = liqThermoSystem.readObject(liqNumb);
            liqThermoSystem.setTemperature(this.Range["B6"].Value2 + 273.15);
            liqThermoSystem.setPressure(this.Range["B7"].Value2);
            ThermodynamicOperations testOps = new ThermodynamicOperations(liqThermoSystem);
            testOps.TPflash();
            liqThermoSystem.display();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

    

    }
}
