﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Security.Principal;
using System.Windows.Forms;
using DatabaseConnection;
using DatabaseConnection.NeqSimDatabaseSetTableAdapters;
using Microsoft.Office.Interop.Excel;
using processSimulation.processEquipment;
using processSimulation.processEquipment.compressor;
using processSimulation.processEquipment.heatExchanger;
using processSimulation.processEquipment.stream;
using processSimulation.processEquipment.util;
using processSimulation.processSystem;
using processSimulation.processSystem.processModules;
using thermo.system;
using thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class SepProcessSheet
    {
        private SystemInterface feedThermoSystem;
        private ProcessSystem operations; // = new processSimulation.processSystem.ProcessSystem();
        private Range statusRange;


        private void Sheet23_Startup(object sender, EventArgs e)
        {
            statusRange = Range["C24"];
            try
            {
                var test = new fluidinfoTableAdapter();
                // NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter test = new NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter test = new NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfo1TableAdapter test = new neqsimdatabaseDataSetTableAdapters.fluidinfo1TableAdapter();

                var userName = WindowsIdentity.GetCurrent().Name;
                userName = userName.Replace("STATOIL-NET\\", "");
                userName = userName.Replace("WIN-NTNU-NO\\", "");
                userName = userName.ToLower();


                var tt = test.GetDataBy(userName);
                var names = new List<string>();
                //names.Add("CPApackage");
                //names.Add(WindowsIdentity.GetCurrent().Name);
                fluidsComboBox.Items.Add("Active fluid");
                foreach (NeqSimDatabaseSet.fluidinfoRow row in tt.Rows)
                {
                    names.Add(row.ID.ToString());
                    fluidsComboBox.Items.Add(row.ID.ToString());
                }

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

        private void Sheet23_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            calcuateButton.Click += calcuateButton_Click;
            fluidsComboBox.SelectedIndexChanged += fluidsComboBox_SelectedIndexChanged;
            linkLabel1.LinkClicked += linkLabel1_LinkClicked;
            linkLabel2.LinkClicked += linkLabel2_LinkClicked;
            linkLabel3.LinkClicked += linkLabel3_LinkClicked;
            Startup += Sheet23_Startup;
            Shutdown += Sheet23_Shutdown;
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

            feedThermoSystem.setTemperature(Range["B7"].Value2 + 273.15);
            feedThermoSystem.setPressure(Range["B6"].Value2);

            statusRange.Value2 = "reading feed fluid...ok";

            statusRange.Value2 = "calculating.....";
            operations = new ProcessSystem();
            var wellStream = new Stream("Well stream", feedThermoSystem);

            wellStream.getThermoSystem().setTotalFlowRate(Range["B5"].Value2, "MSm^3/day");
            var separationModule = new SeparationTrainModule();
            separationModule.addInputStream("feed stream", wellStream);
            separationModule.setSpecification("Second stage pressure", Range["B14"].Value2);
            separationModule.setSpecification("heated oil temperature", 273.15 + Range["B13"].Value2);
            separationModule.setSpecification("Third stage pressure", Range["B16"].Value2);
            separationModule.setSpecification("Gas exit temperature", 273.15 + Range["B12"].Value2);
            separationModule.setSpecification("First stage compressor after cooler temperature",
                273.15 + Range["B17"].Value2);
            separationModule.setSpecification("Export oil temperature", 273.15 + Range["B17"].Value2);

            operations.add(wellStream);
            operations.add(separationModule);
            ((Recycle) operations.getUnit("Resycle")).setTolerance(Range["B19"].Value2);
            ((Compressor) separationModule.getOperations().getUnit("2nd stage recompressor")).setIsentropicEfficiency(
                Range["B20"].Value2 / 100.0);
            ((Compressor) separationModule.getOperations().getUnit("3rd stage recompressor")).setIsentropicEfficiency(
                Range["B20"].Value2 / 100.0);

            operations.run();

            Range["G5"].Value2 =
                separationModule.getOutputStream("gas exit stream").getThermoSystem().getTotalNumberOfMoles() * 8.314 *
                (273.15 + 15.0) / 101325.0 * 3600.0 * 24 / 1.0e6;
            Range["G6"].Value2 =
                separationModule.getOutputStream("oil exit stream").getThermoSystem().getTotalNumberOfMoles() *
                separationModule.getOutputStream("oil exit stream").getThermoSystem().getMolarMass() / separationModule
                    .getOutputStream("oil exit stream").getThermoSystem().getPhase(0).getPhysicalProperties()
                    .getDensity() * 3600.0;
            Range["G7"].Value2 = ((Compressor) separationModule.getOperations().getUnit("3rd stage recompressor"))
                .getPower();
            Range["G8"].Value2 = ((Compressor) separationModule.getOperations().getUnit("2nd stage recompressor"))
                .getPower();
            Range["G9"].Value2 =
                ((Cooler) separationModule.getOperations().getUnit("3rd stage cooler")).getEnergyInput();
            Range["G10"].Value2 = ((Cooler) separationModule.getOperations().getUnit("HP gas cooler")).getEnergyInput();
            Range["G11"].Value2 =
                ((Heater) separationModule.getOperations().getUnit("oil/water heater")).getEnergyInput();
            Range["G12"].Value2 =
                ((Cooler) separationModule.getOperations().getUnit("export oil cooler")).getEnergyInput();

            unitOpsNamesCheckBox.Items.Clear();
            foreach (string name in operations.getAllUnitNames()) unitOpsNamesCheckBox.Items.Add(name);
            //   packageNames = names.ToArray();
            //   fluidListNameComboBox.Items.Add(names.ToList());
            unitOpsNamesCheckBox.SelectedIndex = 0;

            try
            {
                operations.getSystemMechanicalDesign().runDesignCalculation();
                Range["G15"].Value2 = operations.getSystemMechanicalDesign().getTotalNumberOfModules();
                Range["G16"].Value2 = operations.getSystemMechanicalDesign().getTotalPlotSpace();
                Range["G17"].Value2 = operations.getSystemMechanicalDesign().getTotalVolume();
                Range["G18"].Value2 = operations.getSystemMechanicalDesign().getTotalWeight();
                Range["G19"].Value2 = operations.getCostEstimator().getWeightBasedCAPEXEstimate();
            }
            catch (Exception ex)
            {
                Range["G15"].Value2 = "Could not calculate dimensions";
                ex.StackTrace.ToString();
            }

            statusRange.Value2 = "finished.....ok";
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (fluidsComboBox.SelectedItem.ToString().Equals("Active fluid"))
            {
                feedThermoSystem = NeqSimThermoSystem.getThermoSystem();
            }
            else
            {
                var gasNumb = Convert.ToInt32(fluidsComboBox.SelectedItem.ToString());
                feedThermoSystem = feedThermoSystem.readObject(gasNumb);
            }

            feedThermoSystem.setTemperature(Range["B7"].Value2 + 273.15);
            feedThermoSystem.setPressure(Range["B6"].Value2);
            var testOps = new ThermodynamicOperations(feedThermoSystem);
            testOps.TPflash();
            feedThermoSystem.display();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var name = unitOpsNamesCheckBox.SelectedItem.ToString();
            ((ProcessEquipmentBaseClass) operations.getUnit(name)).displayResult();
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var name = unitOpsNamesCheckBox.SelectedItem.ToString();
            ((ProcessEquipmentBaseClass) operations.getUnit(name)).getMechanicalDesign().calcDesign();
            ((ProcessEquipmentBaseClass) operations.getUnit(name)).getMechanicalDesign().displayResults();
        }
    }
}