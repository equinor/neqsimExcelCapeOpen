using System;
using Microsoft.Office.Interop.Excel;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet3
    {
        private void Sheet3_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet3_Shutdown(object sender, EventArgs e)
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
            this.Startup += new System.EventHandler(this.Sheet3_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet3_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var thermoSystem = NeqSimThermoSystem.getThermoSystem(); //.clone();


            var rangeClear = Range["E:L"];
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
                   // thermoSystem.init(0);
                   // thermoSystem.setMultiPhaseCheck(true);
                    //if (thermoSystem.isChemicalSystem()) thermoSystem.setMultiPhaseCheck(false);
                    thermoSystem.setTemperature(r.Value2 + 273.15);
                    thermoSystem.setPressure(Range[textVar].Value2);
                    var ops = new ThermodynamicOperations(thermoSystem);
                    if (thermoSystem.doSolidPhaseCheck())
                        ops.TPSolidflash();
                    else
                        ops.TPflash();

                    var table = thermoSystem.createTable("fluid");
                    var rows = table.Length;
                    var columns = table[1].Length;
                    writeEndCell = writeStartCell + rows;

                    var startCell = Cells[writeStartCell, 5];
                    var endCell = Cells[writeEndCell - 1, columns + 4];
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

        private void button1_Click2(object sender, EventArgs e)
        {
            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            thermoSystem.setTemperature(Range["A2"].Value2 + 273.15);
            thermoSystem.setPressure(Range["B2"].Value2);

            var ops = new ThermodynamicOperations(thermoSystem);
            ops.TPflash();

            // thermoSystem.display();

            var table = thermoSystem.createTable("fluid");
            var rows = table.Length;
            var columns = table[1].Length;

            var startCell = Cells[8, 1];
            var endCell = Cells[rows + 8, columns];
            var writeRange = Range[startCell, endCell];

            //writeRange.Value2 = table;

            var data = new object[rows, columns];
            for (var row = 1; row <= rows; row++)
            for (var column = 1; column <= columns; column++)
                data[row - 1, column - 1] = table[row - 1][column - 1];

            writeRange.Value2 = data;
        }

        private void Sheet3_ActivateEvent()
        {
        }
    }
}