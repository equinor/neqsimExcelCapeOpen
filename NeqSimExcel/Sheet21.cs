using System;
using Microsoft.Office.Interop.Excel;
using neqsim.processSimulation.processEquipment.stream;
using neqsim.processSimulation.processEquipment.valve;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class valveSheet
    {
        private void Sheet21_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet21_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);

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
                var text = (string) r.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    number++;
                    var textVar = "B" + number;
                    var pouttext = "C" + number;
                    var touttext = "D" + number;
                    if (r.Value2 < Range[pouttext].Value2)
                    {
                        Range[touttext].Value2 = "P in lower than P out";
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
                    ThrottlingValve valve_1 = new ThrottlingValve(stream);
                    valve_1.setOutletPressure(Range[pouttext].Value2);
                    stream.run();
                    valve_1.run();
                    Range[touttext].Value2 = valve_1.getOutStream().getThermoSystem().getTemperature() - 273.15;
                    var table = valve_1.getOutStream().getThermoSystem().createTable("fluid");
                    var rows = table.Length;
                    var columns = table[1].Length;
                    writeEndCell = writeStartCell + rows;

                    var startCell = Cells[writeStartCell, 6];
                    var endCell = Cells[writeEndCell - 1, columns + 5];
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