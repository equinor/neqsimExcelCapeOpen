using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet16
    {
        private void Sheet16_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet16_Shutdown(object sender, System.EventArgs e)
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
            this.Startup += new System.EventHandler(this.Sheet16_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet16_Shutdown);

        }

        #endregion

        private void calcuateButton_Click(object sender, EventArgs e)
        {

            double simulationTIme = Range["B4"].Value2;
            double initialVolume = Range["B5"].Value2;
            double boilOffRate = Range["B6"].Value2;
            double pressureShip = Range["B7"].Value2;


            NeqSimThermoSystem.getThermoSystem().setPressure(pressureShip, "bara");
            NeqSimThermoSystem.getThermoSystem().setTemperature(110.0, "K");

            neqsim.fluidMechanics.flowSystem.twoPhaseFlowSystem.shipSystem.LNGship ship = new neqsim.fluidMechanics.flowSystem.twoPhaseFlowSystem.shipSystem.LNGship(NeqSimThermoSystem.getThermoSystem(), initialVolume, boilOffRate / 100.0);
            ship.useStandardVersion("", "2016");
            ship.getStandardISO6976().setEnergyRefT(Convert.ToDouble(energyRefTempComboBox.SelectedItem.ToString()));
            ship.getStandardISO6976().setVolRefT(Convert.ToDouble(isoStandardVolumeRefTempComboBox.SelectedItem.ToString()));
            ship.setEndTime(simulationTIme);
            ship.createSystem();
            ship.init();
            ship.solveSteadyState(0);
            ship.solveTransient(0);
            ship.getResults("temp");
            

            var table = ship.getResultTable();
            var rows = table.Length;
            var columns = table[1].Length;
            int writeStartCell = 1, writeEndCell = 8;

            writeEndCell = writeStartCell + rows;

            var startCell = Cells[writeStartCell, 5];
            var endCell = Cells[writeEndCell - 1, columns + 4];
            var writeRange = Range[startCell, endCell];

            writeStartCell += rows + 3;
            //writeRange.Value2 = table;

            var data = new object[rows, columns];
            table[0][6] = "methane";
            table[0][7] = "ethane";
            table[0][8] = "propane";
            table[0][9] = "i-butane";
            table[0][10] = "n-butane";
            table[0][11] = "i-pentane";
            table[0][12] = "n-pentane";
            table[0][13] = "n-hexane";
            table[0][14] = "nitrogen";

            table[0][15] = "energy";
            table[0][15] = "GCV_mass";

            table[0][17] = "gas_methane";
            table[0][18] = "gas_ethane";
            table[0][19] = "gas_propane";
            table[0][20] = "gas_i-butane";
            table[0][21] = "gas_n-butane";
            table[0][22] = "gas_i-pentane";
            table[0][23] = "gas_n-pentane";
            table[0][24] = "gas_n-hexane";
            table[0][25] = "gas_nitrogen";

            for (var row = 1; row <= rows; row++)
                for (var column = 1; column <= columns; column++)
                    try
                    {
                        String text = table[row - 1][column - 1].Replace(",", ".");
                        text = text.Replace(" ", "");
                        
                        data[row - 1, column - 1] = text;
                    }
                    catch(Exception exe)
                    {
                        data[row - 1, column - 1] = table[row - 1][column - 1];
                    }

            writeRange.Value2 = data;
        }
    }
}
