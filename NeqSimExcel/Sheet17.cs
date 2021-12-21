using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Security.Principal;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using neqsim.thermo.system;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet17
    {
        private SystemInterface gasThermoSystem, liqThermoSystem;
        private Range statusRange;

        private void Sheet17_Startup(object sender, EventArgs e)
        {
            statusRange = Range["D14"];
          

        }

        private void Sheet17_Shutdown(object sender, EventArgs e)
        {
        }

        private void ActivateWorksheet()
        {
            fluidListNameComboBoxGas.Items.Clear();
            fluidListNameComboBoxLiquid.Items.Clear();
            try
            {
                var names = new List<string>();
                var filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                var fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";

                var d = new DirectoryInfo(fullPath);
                var Files = d.GetFiles("*.neqsim");
                foreach (var file in Files)
                {
                    names.Add(file.Name.Replace(".neqsim", ""));
                    fluidListNameComboBoxGas.Items.Add(file.Name.Replace(".neqsim", ""));
                    fluidListNameComboBoxLiquid.Items.Add(file.Name.Replace(".neqsim", ""));
                }

                fluidListNameComboBoxGas.SelectedIndex = 0;
                fluidListNameComboBoxLiquid.SelectedIndex = 0;

                var gasNumb = Convert.ToInt32(fluidListNameComboBoxGas.SelectedItem.ToString());
                gasThermoSystem = NeqSimThermoSystem.getThermoSystem();
                gasThermoSystem = gasThermoSystem.readObject(gasNumb);

                var liqNumb = Convert.ToInt32(fluidListNameComboBoxLiquid.SelectedItem.ToString());
                liqThermoSystem = NeqSimThermoSystem.getThermoSystem();
                liqThermoSystem = liqThermoSystem.readObject(liqNumb);
            }
            catch (Exception excet)
            {
                Console.WriteLine("Error " + excet.Message);
            }
            
            fluidListNameComboBoxGas.SelectedIndex = 0;
            fluidListNameComboBoxLiquid.SelectedIndex = 0;
        }

        #region VSTO Designer generated code

            /// <summary>
            ///     Required method for Designer support - do not modify
            ///     the contents of this method with the code editor.
            /// </summary>
            private void InternalStartup()
        {
            this.calculateButton.Click += new System.EventHandler(this.calculateButton_Click);
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_Click);
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.ActivateWorksheet);
            this.Startup += new System.EventHandler(this.Sheet17_Startup);

        }

        #endregion

        private void fluidListNameComboBoxGas_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void calculateButton_Click(object sender, EventArgs e)
        {
            statusRange.Font.Color = ColorTranslator.ToOle(Color.Blue);
            Range["F2", "H100"].Clear();

            statusRange.Value2 = "reading fluids...";

            string name = fluidListNameComboBoxGas.SelectedItem.ToString();
            var filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";
            var filename = fullPath + "\\" + name + ".neqsim";
            gasThermoSystem = NeqSimThermoSystem.getThermoSystem();
            gasThermoSystem = gasThermoSystem.readObjectFromFile(filename, filename);
            gasThermoSystem.setTemperature(Range["B6"].Value2 + 273.15);
            gasThermoSystem.setPressure(Range["B7"].Value2);
            var gasOps = new ThermodynamicOperations(gasThermoSystem);
            gasOps.TPflash();
            if (!gasThermoSystem.hasPhaseType("gas"))
            {
                statusRange.Font.Color = ColorTranslator.ToOle(Color.Red);
                statusRange.Value2 = "no gas at given temperature and pressure...check gas";
                return;
            }

            statusRange.Value2 = "reading gas...ok";

            name = fluidListNameComboBoxLiquid.SelectedItem.ToString();
            filename = fullPath + "\\" + name + ".neqsim";
            liqThermoSystem = NeqSimThermoSystem.getThermoSystem();
            liqThermoSystem = liqThermoSystem.readObjectFromFile(filename, filename);
            liqThermoSystem.setTotalFlowRate(Range["B11"].Value2, "kg/sec");
            liqThermoSystem.init(0);
            liqThermoSystem.init(1);
            liqThermoSystem.setMultiPhaseCheck(true);
            liqThermoSystem.setTemperature(Range["B6"].Value2 + 273.15);
            liqThermoSystem.setPressure(Range["B7"].Value2);
            var liqOps = new ThermodynamicOperations(liqThermoSystem);
            liqOps.TPflash();
            if (!(liqThermoSystem.hasPhaseType("aqueous") || liqThermoSystem.hasPhaseType("liquid")))
            {
                statusRange.Font.Color = ColorTranslator.ToOle(Color.Red);
                statusRange.Value2 = "no liquid phase at given temperature and pressure...check liquid";
                return;
            }


            statusRange.Value2 = "reading liquid...ok";
            if (!liqThermoSystem.getPhase(0).hasComponent("water"))
            {
                statusRange.Font.Color = ColorTranslator.ToOle(Color.Red);
                statusRange.Value2 = "liquid phase must contain water... stopping calcuations";
                return;
            }

            var testOps = new ThermodynamicOperations(liqThermoSystem);


            var oldTime = 0.0;
            var range = Range["E2", "E100"];
            var number = 0;

            statusRange.Value2 = "calculations running.....";
            foreach (Range r in range.Cells)
            {
                var text = (string) r.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    double time = Range["E" + (number + 2)].Value2;
                    gasThermoSystem.setTotalFlowRate(Range["B5"].Value2 * (time - oldTime) / 24.0 * 3600, "MSm3/day");
                    oldTime = time;
                    gasThermoSystem.init(1);
                    liqThermoSystem.addFluid(gasThermoSystem);
                    testOps.TPflash();
                    var phaseN = liqThermoSystem.getNumberOfPhases();
                    if (phaseN > 1)
                    {
                        Range["F" + (number + 2)].Value2 =
                            liqThermoSystem.getPhase(phaseN - 1).getWtFrac(liqThermoSystem.getPhase(phaseN - 1)
                                .getComponent("water").getComponentNumber()) * 100.0;
                        Range["H" + (number + 2)].Value2 =
                            liqThermoSystem.getPhase(0).getComponent("water").getx() * 1e6;
                        Range["G" + (number + 2)].Value2 =
                            liqThermoSystem.getPhase(phaseN - 1).getNumberOfMolesInPhase() *
                            liqThermoSystem.getPhase(phaseN - 1).getMolarMass();
                        liqThermoSystem = liqThermoSystem.phaseToSystem(phaseN - 1);
                    }

                    testOps = new ThermodynamicOperations(liqThermoSystem);
                    // System.out.println("nuber of moles " + testSystem.getNumberOfMoles() + " moleFrac MEG " + testSystem.getPhase(0).getComponent("MEG").getx());
                    number++;
                }
            }

            statusRange.Font.Color = ColorTranslator.ToOle(Color.Green);
            statusRange.Value2 = "finished calcuatons";
        }


        private void linkLabel1_Click(object sender, EventArgs e)
        {
            string name = fluidListNameComboBoxGas.SelectedItem.ToString();
            var filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";
            var filename = fullPath + "\\" + name + ".neqsim";
            gasThermoSystem = NeqSimThermoSystem.getThermoSystem();
            gasThermoSystem = gasThermoSystem.readObjectFromFile(filename, filename);
            gasThermoSystem.setTemperature(Range["B6"].Value2 + 273.15);
            gasThermoSystem.setPressure(Range["B7"].Value2);
            var testOps = new ThermodynamicOperations(gasThermoSystem);
            testOps.TPflash();
            gasThermoSystem.display();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string name = fluidListNameComboBoxLiquid.SelectedItem.ToString();
            var filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";
            var filename = fullPath + "\\" + name + ".neqsim";
            liqThermoSystem = NeqSimThermoSystem.getThermoSystem();
            liqThermoSystem = liqThermoSystem.readObjectFromFile(filename, filename);
            liqThermoSystem.setTemperature(Range["B6"].Value2 + 273.15);
            liqThermoSystem.setPressure(Range["B7"].Value2);
            var testOps = new ThermodynamicOperations(liqThermoSystem);
            testOps.TPflash();
            liqThermoSystem.display();
        }

    }
}