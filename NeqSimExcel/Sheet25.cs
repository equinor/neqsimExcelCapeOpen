using System;
using System.Collections.Generic;
using System.Drawing;
using System.Security.Principal;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using neqsim.processSimulation.processEquipment;
using neqsim.processSimulation.processEquipment.compressor;
using neqsim.processSimulation.processEquipment.heatExchanger;
using neqsim.processSimulation.processEquipment.pump;
using neqsim.processSimulation.processEquipment.stream;
using neqsim.processSimulation.processEquipment.util;
using neqsim.processSimulation.processSystem;
using neqsim.processSimulation.processSystem.processModules;
using neqsim.thermo.system;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;


namespace NeqSimExcel
{
    public partial class Sheet25
    {
        private SystemInterface feedThermoSystem;
        private ProcessSystem operations; // = new processSimulation.processSystem.ProcessSystem();
        private Range statusRange;

        private void Sheet25_Startup(object sender, EventArgs e)
        {
            statusRange = Range["C28"];
            
        }

        private void Sheet25_Shutdown(object sender, EventArgs e)
        {
        }

        private void ActivateWorkSheet()
        {
            fluidsComboBox.Items.Clear();
            try
            {
                // NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter test = new NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter test = new NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfo1TableAdapter test = new neqsimdatabaseDataSetTableAdapters.fluidinfo1TableAdapter();

                var userName = WindowsIdentity.GetCurrent().Name;
                userName = userName.Replace("STATOIL-NET\\", "");
                userName = userName.Replace("WIN-NTNU-NO\\", "");
                userName = userName.ToLower();


                var names = new List<string>();
                //names.Add("CPApackage");
                //names.Add(WindowsIdentity.GetCurrent().Name);
                fluidsComboBox.Items.Add("Active fluid");
              

                //   packageNames = names.ToArray();
                //   fluidListNameComboBox.Items.Add(names.ToList());
                fluidsComboBox.SelectedIndex = 0;

                var feedNumb = Convert.ToInt32(fluidsComboBox.SelectedItem.ToString());
                feedThermoSystem = NeqSimThermoSystem.getThermoSystem();
                feedThermoSystem = feedThermoSystem.readObject(feedNumb);
            }
            catch (Exception excet)
            {
                Console.WriteLine("Error " + excet.Message);
            }
        }

        #region VSTO Designer generated code

            /// <summary>
            ///     Required method for Designer support - do not modify
            ///     the contents of this method with the code editor.
            /// </summary>
            private void InternalStartup()
        {
            this.calcuateButton.Click += new System.EventHandler(this.calcuateButton_Click);
            this.fluidsComboBox.SelectedIndexChanged += new System.EventHandler(this.fluidsComboBox_SelectedIndexChanged);
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked_1);
            this.unitOpsNamesCheckBox.SelectedIndexChanged += new System.EventHandler(this.unitOpsNamesCheckBox_SelectedIndexChanged);
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked_1);
            this.glycolTypeComboBox.SelectedIndexChanged += new System.EventHandler(this.glycolTypeComboBox_SelectedIndexChanged);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.ActivateWorkSheet);
            this.Startup += new System.EventHandler(this.Sheet25_Startup);

        }

        #endregion

        private void fluidsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void calcuateButton_Click(object sender, EventArgs e)
        {
            var rangeClear = Range["G1", "G100"];
            rangeClear.Clear();
            statusRange.Font.Color = ColorTranslator.ToOle(Color.Blue);

            if (fluidsComboBox.SelectedItem.ToString().Equals("Active fluid"))
            {
                feedThermoSystem = NeqSimThermoSystem.getThermoSystem();
            }
            else
            {
                statusRange.Value2 = "reading fluids...";
                var gasNumb = Convert.ToInt32(fluidsComboBox.SelectedItem.ToString());
                feedThermoSystem = NeqSimThermoSystem.getThermoSystem();
                feedThermoSystem = feedThermoSystem.readObject(gasNumb);
            }

            var glycolName = glycolTypeComboBox.SelectedItem.ToString();

            if (!feedThermoSystem.getPhase(0).hasComponent("water")) feedThermoSystem.addComponent("water", 1e-10);
            if (!feedThermoSystem.getPhase(0).hasComponent(glycolName))
                feedThermoSystem.addComponent(glycolName, 1e-10);

            feedThermoSystem.createDatabase(true);
            feedThermoSystem.setMixingRule(10);
            feedThermoSystem.setMultiPhaseCheck(true);
            feedThermoSystem.init(0);
            feedThermoSystem = feedThermoSystem.autoSelectModel();
            //feedThermoSystem = feedThermoSystem.autoSelectModel();
            feedThermoSystem.setTemperature(Range["B7"].Value2 + 273.15);
            feedThermoSystem.setPressure(Range["B6"].Value2);

            SystemInterface glycolTestSystem = (SystemInterface) feedThermoSystem.clone();
            glycolTestSystem.removeMoles();
            glycolTestSystem.addComponent("water", 100 - Range["B21"].Value2, "kg/hr");
            glycolTestSystem.addComponent(glycolName, Range["B21"].Value2, "kg/hr");
            glycolTestSystem.init(0);
            glycolTestSystem.setTemperature(Range["B22"].Value2 + 273.15);

            var glycolFeedStream = new Stream("Glycol feed stream", glycolTestSystem);
            glycolFeedStream.getThermoSystem().setTotalFlowRate(Range["B20"].Value2 * 1000.0, "kg/hr");

            statusRange.Value2 = "reading feed fluid...ok";

            statusRange.Value2 = "calculating.....";
            operations = new ProcessSystem();
            var wellStream = new Stream("Well stream", feedThermoSystem);

            wellStream.getThermoSystem().setTotalFlowRate(Range["B5"].Value2, "MSm^3/day");
            var separationModule = new MixerGasProcessingModule("mixer module");
            separationModule.addInputStream("feed stream", wellStream);
            separationModule.addInputStream("glycol feed stream", glycolFeedStream);
            separationModule.setSpecification("inlet separation temperature", Range["B12"].Value2);
            separationModule.setSpecification("gas scrubber temperature", Range["B13"].Value2);
            separationModule.setSpecification("glycol scrubber temperature", Range["B14"].Value2);
            separationModule.setSpecification("first stage out pressure", Range["B15"].Value2);
            separationModule.setSpecification("second stage out pressure", Range["B18"].Value2);
            separationModule.setSpecification("export gas temperature", Range["B17"].Value2);
            separationModule.setSpecification("liquid pump out pressure", Range["B16"].Value2);

            operations.add(wellStream);
            operations.add(glycolFeedStream);
            operations.add(separationModule);

            ((Recycle) operations.getUnit("Resycle")).setTolerance(Range["B23"].Value2);
            ((Compressor) separationModule.getOperations().getUnit("1st stage compressor")).setIsentropicEfficiency(
                Range["B24"].Value2 / 100.0);
            ((Compressor) separationModule.getOperations().getUnit("2nd stage compressor")).setIsentropicEfficiency(
                Range["B24"].Value2 / 100.0);

            operations.run();

            Range["G5"].Value2 =
                separationModule.getOutputStream("gas exit stream").getThermoSystem().getTotalNumberOfMoles() * 8.314 *
                (273.15 + 15.0) / 101325.0 * 3600.0 * 24 / 1.0e6;
            Range["G6"].Value2 =
                separationModule.getOutputStream("oil exit stream").getThermoSystem().getTotalNumberOfMoles() *
                separationModule.getOutputStream("oil exit stream").getThermoSystem().getMolarMass() / separationModule
                    .getOutputStream("oil exit stream").getThermoSystem().getPhase(0).getPhysicalProperties()
                    .getDensity() * 3600.0;
            Range["G7"].Value2 =
                ((Compressor) separationModule.getOperations().getUnit("1st stage compressor")).getPower();
            Range["G8"].Value2 =
                ((Compressor) separationModule.getOperations().getUnit("2nd stage compressor")).getPower();
            Range["G9"].Value2 = ((Cooler) separationModule.getOperations().getUnit("separator gas cooler"))
                .getEnergyInput();
            Range["G10"].Value2 = ((Cooler) separationModule.getOperations().getUnit("glycol mixer after cooler"))
                .getEnergyInput();
            Range["G11"].Value2 = ((Cooler) separationModule.getOperations().getUnit("second stage after cooler"))
                .getEnergyInput();
            Range["G12"].Value2 = ((Cooler) separationModule.getOperations().getUnit("inlet well stream cooler"))
                .getEnergyInput();
            Range["G13"].Value2 = separationModule.getOutputStream("gas exit stream").getThermoSystem().getPhase(0)
                                      .getComponent("water").getx() * 1e6;
            Range["G14"].Value2 = ((Pump) separationModule.getOperations().getUnit("liquid pump")).getPower();

            unitOpsNamesCheckBox.Items.Clear();
            foreach (string name in operations.getAllUnitNames()) unitOpsNamesCheckBox.Items.Add(name);
            //   packageNames = names.ToArray();
            //   fluidListNameComboBox.Items.Add(names.ToList());
            unitOpsNamesCheckBox.SelectedIndex = 0;


            statusRange.Value2 = "finished.....ok";
        }

        private void linkLabel1_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (fluidsComboBox.SelectedItem.ToString().Equals("Active fluid"))
            {
                feedThermoSystem = NeqSimThermoSystem.getThermoSystem();
            }
            else
            {
                var gasNumb = Convert.ToInt32(fluidsComboBox.SelectedItem.ToString());
                feedThermoSystem = NeqSimThermoSystem.getThermoSystem();
            }

            feedThermoSystem.setTemperature(Range["B7"].Value2 + 273.15);
            feedThermoSystem.setPressure(Range["B6"].Value2);
            var testOps = new ThermodynamicOperations(feedThermoSystem);
            testOps.TPflash();
            feedThermoSystem.display();
        }

        private void linkLabel2_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var name = unitOpsNamesCheckBox.SelectedItem.ToString();
            ((ProcessEquipmentBaseClass) operations.getUnit(name)).displayResult();
        }

        private void unitOpsNamesCheckBox_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void glycolTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
    }
}