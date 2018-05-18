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

namespace NeqSimExcel
{
    public partial class Sheet3
    {
        private void Sheet3_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet3_Shutdown(object sender, System.EventArgs e){
        
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.Sheet3_ActivateEvent);
            this.Startup += new System.EventHandler(this.Sheet3_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet3_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            SystemInterface thermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem();//.clone();


            Excel.Range rangeClear = this.Range["E:L"];
            rangeClear.Clear();
            Excel.Range range = this.Range["A2", "A100"];
            int number = 1;

            int writeStartCell = 1, writeEndCell=8;
            
            foreach (Excel.Range r in range.Cells)
            {
                string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    number++;
                    string textVar = "B" + number.ToString();
                    thermoSystem.init(0);
                    thermoSystem.setMultiPhaseCheck(true);
                    if(thermoSystem.isChemicalSystem()) thermoSystem.setMultiPhaseCheck(false);
                    thermoSystem.setTemperature(r.Value2+273.15);
                    thermoSystem.setPressure(this.Range[textVar].Value2);
                    ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
                    if (thermoSystem.doSolidPhaseCheck())
                    {
                        ops.TPSolidflash();
                    }
                    else
                    {
                        ops.TPflash();
                    }

                var table = thermoSystem.createTable("fluid");
                int rows = table.Length;
                int columns = table[1].Length;
                writeEndCell = writeStartCell+rows;

                var startCell = Cells[writeStartCell, 5];
                var endCell = Cells[writeEndCell-1, columns+4];
                var writeRange = this.Range[startCell, endCell];

                writeStartCell += rows+3;
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

        private void button1_Click2(object sender, EventArgs e)
        {

            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            thermoSystem.setTemperature(this.Range["A2"].Value2 + 273.15);
            thermoSystem.setPressure(this.Range["B2"].Value2);

            ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
            ops.TPflash();

            // thermoSystem.display();

            var table = thermoSystem.createTable("fluid");
            int rows = table.Length;
            int columns = table[1].Length;

            var startCell = Cells[8, 1];
            var endCell = Cells[rows + 8, columns];
            var writeRange = this.Range[startCell, endCell];

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

        private void Sheet3_ActivateEvent()
        {
        }

    }
}
