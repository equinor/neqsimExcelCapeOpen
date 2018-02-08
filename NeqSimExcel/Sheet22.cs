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
using processSimulation.processEquipment.compressor;
using processSimulation.processEquipment.stream;

namespace NeqSimExcel
{
    public partial class compressorSheet
    {
        private void Sheet22_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet22_Shutdown(object sender, System.EventArgs e)
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
            this.Startup += new System.EventHandler(this.Sheet22_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet22_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            SystemInterface thermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem();//.clone();


            Excel.Range rangeClear = this.Range["E2", "L100"];
            rangeClear.Clear();
            Excel.Range range = this.Range["A2", "A100"];
            int number = 1;

            int writeStartCell = 1, writeEndCell = 8;

            foreach (Excel.Range r in range.Cells)
            {
                string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    number++;
                    string textVar = "B" + number.ToString();
                    string pouttext = "C" + number.ToString();
                    string efficiencytest = "D" + number.ToString();
                    string  touttext= "E" + number.ToString();
                    string powertext = "F" + number.ToString();
                    if (r.Value2 > this.Range[pouttext].Value2)
                    {
                        this.Range[touttext].Value2 = "P out lower than P in";
                        return;
                    }
                    thermoSystem.init(0);

                    thermoSystem.setPressure(r.Value2);
                    thermoSystem.setTemperature(this.Range[textVar].Value2 + 273.15);
                    ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
                    if (thermoSystem.doSolidPhaseCheck())
                    {
                        ops.TPSolidflash();
                    }
                    else
                    {
                        ops.TPflash();
                    }

                    Stream stream = new Stream("Stream1", thermoSystem);
                    Compressor compressor = new Compressor(stream);
                    compressor.setOutletPressure(this.Range[pouttext].Value2);
                    compressor.setIsentropicEfficiency(this.Range[efficiencytest].Value2/100.0);
                    stream.run();
                    compressor.run();
                    this.Range[touttext].Value2 = compressor.getOutStream().getThermoSystem().getTemperature() - 273.15;
                    this.Range[powertext].Value2 = compressor.getEnergy()/(compressor.getOutStream().getThermoSystem().getNumberOfMoles()*compressor.getOutStream().getThermoSystem().getMolarMass());
                    var table = compressor.getOutStream().getThermoSystem().createTable("fluid");
                    int rows = table.Length;
                    int columns = table[1].Length;
                    writeEndCell = writeStartCell + rows;

                    var startCell = Cells[writeStartCell, 8];
                    var endCell = Cells[writeEndCell - 1, columns + 7];
                    var writeRange = this.Range[startCell, endCell];

                    writeStartCell += rows + 3;
                    //writeRange.Value2 = table;

                    var data = new object[rows, columns];
                    for (var row = 1; row <= rows; row++)
                    {
                        for (var column = 1; column <= columns; column++)
                        {
                            data[row - 1, column - 1] = table[row - 1][column - 1];
                        }
                    }

                    writeRange.Value2 = data;
                }
            }

        }

    }
}
