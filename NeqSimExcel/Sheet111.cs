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
    public partial class Sheet11
    {
        private void Sheet11_Startup(object sender, System.EventArgs e)
        {
            componentComboBox.Items.Clear();

            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                string[] names = thermoSystem.getComponentNames();

                foreach (string name in names)
                {
                    componentComboBox.Items.Add(name);
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

        }

        private void Sheet11_Shutdown(object sender, System.EventArgs e)
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
            this.componentComboBox.Click += new System.EventHandler(this.componentComboBox_SelectedIndexChanged);
            this.Startup += new System.EventHandler(this.Sheet11_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet11_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            SystemInterface thermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem();//.clone();
            thermoSystem.setTemperature(thermoSystem.getPhase(0).getComponent(componentComboBox.SelectedItem.ToString()).getTriplePointTemperature());
            thermoSystem.setSolidPhaseCheck(false);
            thermoSystem.setSolidPhaseCheck(componentComboBox.SelectedItem.ToString());

            Excel.Range rangeClear = this.Range["E2", "L100"];
            rangeClear.Clear();
            Excel.Range range = this.Range["A2", "A100"];
            int number = 2;

            int writeStartCell = 1, writeEndCell = 8;

            foreach (Excel.Range r in range.Cells)
            {
                string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    string textVar = "A" + number.ToString();
                    string temperatureVar = "B" + number.ToString();
                    number++;
                    thermoSystem.setPressure(this.Range[textVar].Value2);
                    ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
                    ops.freezingPointTemperatureFlash();
                    this.Range[temperatureVar].Value2 = thermoSystem.getTemperature()-273.15;
                    var table = thermoSystem.createTable("fluid");
                    int rows = table.Length;
                    int columns = table[1].Length;
                    writeEndCell = writeStartCell + rows;

                    var startCell = Cells[writeStartCell, 5];
                    var endCell = Cells[writeEndCell - 1, columns + 4];
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
        private void button2_Click(object sender, EventArgs e)
        {
           componentComboBox.Items.Clear();

            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                string[] names = thermoSystem.getComponentNames();

                foreach (string name in names)
                {
                    componentComboBox.Items.Add(name);
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void componentComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            componentComboBox.Items.Clear();

            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                string[] names = thermoSystem.getComponentNames();

                foreach (string name in names)
                {
                    componentComboBox.Items.Add(name);
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }
   }

}
