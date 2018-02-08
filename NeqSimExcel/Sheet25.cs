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
using processSimulation.processEquipment.stream;
using processSimulation.processEquipment.compressor;
using processSimulation.processEquipment.pump;
using processSimulation.processEquipment.heatExchanger;
using processSimulation.processSystem.processModules;


namespace NeqSimExcel
{
    public partial class Sheet25
    {
        SystemInterface feedThermoSystem;
        Excel.Range statusRange = null;
        processSimulation.processSystem.ProcessSystem operations;// = new processSimulation.processSystem.ProcessSystem();

        private void Sheet25_Startup(object sender, System.EventArgs e)
        {
            statusRange = this.Range["C28"];
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
                fluidsComboBox.Items.Add("Active fluid");
                foreach (DatabaseConnection.NeqSimDatabaseSet.fluidinfoRow row in tt.Rows)
                {
                    names.Add(row.ID.ToString());
                    fluidsComboBox.Items.Add(row.ID.ToString());
                }
                //   packageNames = names.ToArray();
                //   fluidListNameComboBox.Items.Add(names.ToList());
                fluidsComboBox.SelectedIndex = 0;

                int feedNumb = Convert.ToInt32((String)fluidsComboBox.SelectedItem.ToString());
                feedThermoSystem = NeqSimThermoSystem.getThermoSystem();
                feedThermoSystem = feedThermoSystem.readObject(feedNumb);
             
            }
            catch (Exception excet)
            {
                Console.WriteLine("Error " + excet.Message);
            }
        }

        private void Sheet25_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.calcuateButton.Click += new System.EventHandler(this.calcuateButton_Click);
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked_1);
            this.unitOpsNamesCheckBox.SelectedIndexChanged += new System.EventHandler(this.unitOpsNamesCheckBox_SelectedIndexChanged);
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked_1);
            this.glycolTypeComboBox.SelectedIndexChanged += new System.EventHandler(this.glycolTypeComboBox_SelectedIndexChanged);
            this.Startup += new System.EventHandler(this.Sheet25_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet25_Shutdown);

        }

        #endregion

        private void fluidsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void calcuateButton_Click(object sender, EventArgs e)
        {
            Excel.Range rangeClear = this.Range["G1", "G100"];
            rangeClear.Clear();
            statusRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

            if (fluidsComboBox.SelectedItem.ToString().Equals("Active fluid"))
            {
                feedThermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem();
            }
            else
            {
                statusRange.Value2 = "reading fluids...";
                int gasNumb = Convert.ToInt32((String)fluidsComboBox.SelectedItem.ToString());
                feedThermoSystem = NeqSimThermoSystem.getThermoSystem();
                feedThermoSystem = feedThermoSystem.readObject(gasNumb);
            }
            
            string glycolName = glycolTypeComboBox.SelectedItem.ToString();
          
            if(!feedThermoSystem.getPhase(0).hasComponent("water")) feedThermoSystem.addComponent("water", 1e-10);
            if (!feedThermoSystem.getPhase(0).hasComponent(glycolName)) feedThermoSystem.addComponent(glycolName, 1e-10);

            feedThermoSystem.createDatabase(true);
            feedThermoSystem.setMixingRule(10);
            feedThermoSystem.setMultiPhaseCheck(true);
            feedThermoSystem.init(0);
            feedThermoSystem = feedThermoSystem.autoSelectModel();
            //feedThermoSystem = feedThermoSystem.autoSelectModel();
            feedThermoSystem.setTemperature(this.Range["B7"].Value2 + 273.15);
            feedThermoSystem.setPressure(this.Range["B6"].Value2);

            thermo.system.SystemInterface glycolTestSystem = (SystemInterface) feedThermoSystem.clone();
            glycolTestSystem.removeMoles();
            glycolTestSystem.addComponent("water", 100 - this.Range["B21"].Value2, "kg/hr");
            glycolTestSystem.addComponent(glycolName, this.Range["B21"].Value2, "kg/hr");
            glycolTestSystem.init(0);
            glycolTestSystem.setTemperature(this.Range["B22"].Value2+273.15);
            
            Stream glycolFeedStream = new Stream("Glycol feed stream", glycolTestSystem);
            glycolFeedStream.getThermoSystem().setTotalFlowRate(this.Range["B20"].Value2*1000.0, "kg/hr");

            statusRange.Value2 = "reading feed fluid...ok";

            statusRange.Value2 = "calculating.....";
            operations = new processSimulation.processSystem.ProcessSystem();
            Stream wellStream = new Stream("Well stream", feedThermoSystem);

            wellStream.getThermoSystem().setTotalFlowRate(this.Range["B5"].Value2, "MSm^3/day");
            G2PGasProcessingModule separationModule = new G2PGasProcessingModule();
            separationModule.addInputStream("feed stream", wellStream);
            separationModule.addInputStream("glycol feed stream", glycolFeedStream);
            separationModule.setSpecification("inlet separation temperature", this.Range["B12"].Value2);
            separationModule.setSpecification("gas scrubber temperature", this.Range["B13"].Value2);
            separationModule.setSpecification("glycol scrubber temperature", this.Range["B14"].Value2);
            separationModule.setSpecification("first stage out pressure", this.Range["B15"].Value2);
            separationModule.setSpecification("second stage out pressure", this.Range["B18"].Value2);
            separationModule.setSpecification("export gas temperature", this.Range["B17"].Value2);
            separationModule.setSpecification("liquid pump out pressure", this.Range["B16"].Value2);

            operations.add(wellStream);
            operations.add(glycolFeedStream);
            operations.add(separationModule);

            ((processSimulation.processEquipment.util.Recycle)operations.getUnit("Resycle")).setTolerance(this.Range["B23"].Value2);
            ((Compressor)separationModule.getOperations().getUnit("1st stage compressor")).setIsentropicEfficiency(this.Range["B24"].Value2/100.0);
            ((Compressor)separationModule.getOperations().getUnit("2nd stage compressor")).setIsentropicEfficiency(this.Range["B24"].Value2 / 100.0);
          
            operations.run();

            this.Range["G5"].Value2 = separationModule.getOutputStream("gas exit stream").getThermoSystem().getTotalNumberOfMoles() * 8.314 * (273.15 + 15.0) / 101325.0 * 3600.0 * 24 / 1.0e6;
            this.Range["G6"].Value2 = separationModule.getOutputStream("oil exit stream").getThermoSystem().getTotalNumberOfMoles() * separationModule.getOutputStream("oil exit stream").getThermoSystem().getMolarMass() / separationModule.getOutputStream("oil exit stream").getThermoSystem().getPhase(0).getPhysicalProperties().getDensity() * 3600.0;
            this.Range["G7"].Value2 = ((Compressor)separationModule.getOperations().getUnit("1st stage compressor")).getPower();
            this.Range["G8"].Value2 = ((Compressor)separationModule.getOperations().getUnit("2nd stage compressor")).getPower();
            this.Range["G9"].Value2 = ((Cooler)separationModule.getOperations().getUnit("separator gas cooler")).getEnergyInput();
            this.Range["G10"].Value2 = ((Cooler)separationModule.getOperations().getUnit("glycol mixer after cooler")).getEnergyInput();
            this.Range["G11"].Value2 = ((Cooler)separationModule.getOperations().getUnit("second stage after cooler")).getEnergyInput();
            this.Range["G12"].Value2 = ((Cooler)separationModule.getOperations().getUnit("inlet well stream cooler")).getEnergyInput();
            this.Range["G13"].Value2 = separationModule.getOutputStream("gas exit stream").getThermoSystem().getPhase(0).getComponent("water").getx() * 1e6;
            this.Range["G14"].Value2 = ((Pump)separationModule.getOperations().getUnit("liquid pump")).getPower();
          
            unitOpsNamesCheckBox.Items.Clear();
            foreach (String name in operations.getAllUnitNames())
            {
                unitOpsNamesCheckBox.Items.Add(name);
            }
            //   packageNames = names.ToArray();
            //   fluidListNameComboBox.Items.Add(names.ToList());
            unitOpsNamesCheckBox.SelectedIndex = 0;

            
            
            
            statusRange.Value2 = "finished.....ok";
        }

        private void linkLabel1_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (fluidsComboBox.SelectedItem.ToString().Equals("Active fluid"))
            {
                feedThermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem();
            }
            else
            {
                int gasNumb = Convert.ToInt32((String)fluidsComboBox.SelectedItem.ToString());
                feedThermoSystem = NeqSimThermoSystem.getThermoSystem();
            }
             feedThermoSystem.setTemperature(this.Range["B7"].Value2 + 273.15);
            feedThermoSystem.setPressure(this.Range["B6"].Value2);
            ThermodynamicOperations testOps = new ThermodynamicOperations(feedThermoSystem);
            testOps.TPflash();
            feedThermoSystem.display();
        }

        private void linkLabel2_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            String name = (String)unitOpsNamesCheckBox.SelectedItem.ToString();
            ((processSimulation.processEquipment.ProcessEquipmentBaseClass)operations.getUnit(name)).displayResult();
            
        }

        private void unitOpsNamesCheckBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void glycolTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
