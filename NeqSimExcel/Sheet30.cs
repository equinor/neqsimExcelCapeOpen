using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using neqsim.processSimulation.processEquipment.expander;
using neqsim.processSimulation.processEquipment.stream;
using neqsim.thermodynamicOperations;

namespace NeqSimExcel
{
    public partial class Sheet30
    {
        private void Sheet30_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet30_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.Startup += new System.EventHandler(this.Sheet30_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet30_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var thermoSystem = NeqSimThermoSystem.getThermoSystem(); //.clone();


            var rangeClear = Range["E2", "L100"];
            rangeClear.Clear();
            var range = Range["A2", "A100"];
            var number = 1;

            int writeStartCell = 1, writeEndCell = 8;

            foreach (Range r in range.Cells)
            {
                var text = (string)r.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    number++;
                    var textVar = "B" + number;
                    var pouttext = "C" + number;
                    var efficiencytest = "D" + number;
                    var touttext = "E" + number;
                    var powertext = "F" + number;
                    if (r.Value2 < Range[pouttext].Value2)
                    {
                        Range[touttext].Value2 = "P out higher than P in";
                        return;
                    }

                    thermoSystem.init(0);

                    thermoSystem.setPressure(r.Value2);
                    thermoSystem.setTemperature(Range[textVar].Value2 + 273.15);
                    var ops = new ThermodynamicOperations(thermoSystem);
                    if (thermoSystem.doSolidPhaseCheck())
                        ops.TPSolidflash();
                    else
                        ops.TPflash();

                    Stream stream = new Stream("Stream1", thermoSystem);
                    Expander compressor = new Expander("exp1", stream);
                    compressor.setOutletPressure(Range[pouttext].Value2);
                    compressor.setIsentropicEfficiency(Range[efficiencytest].Value2 / 100.0);
                    stream.run();
                    compressor.run();
                    Range[touttext].Value2 = compressor.getOutStream().getThermoSystem().getTemperature() - 273.15;
                    Range[powertext].Value2 = compressor.getEnergy() /
                                              (compressor.getOutStream().getThermoSystem().getNumberOfMoles() *
                                               compressor.getOutStream().getThermoSystem().getMolarMass());
                    var table = compressor.getOutStream().getThermoSystem().createTable("fluid");
                    var rows = table.Length;
                    var columns = table[1].Length;
                    writeEndCell = writeStartCell + rows;

                    var startCell = Cells[writeStartCell, 8];
                    var endCell = Cells[writeEndCell - 1, columns + 7];
                    var writeRange = Range[startCell, endCell];

                    writeStartCell += rows + 3;
                    //writeRange.Value2 = table;

                    var data = new object[rows, columns];
                    for (var row = 1; row <= rows; row++)
                        for (var column = 1; column <= columns; column++)
                            data[row - 1, column - 1] = table[row - 1][column - 1];

                    writeRange.Value2 = data;
                }
            }
        }
    }
}
