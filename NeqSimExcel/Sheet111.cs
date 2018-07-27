using System;
using Microsoft.Office.Interop.Excel;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet11
    {
        private void Sheet11_Startup(object sender, EventArgs e)
        {
           
        }

        private void ActivateComponentCombobox()
        {
            componentComboBox.Items.Clear();

            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                var names = thermoSystem.getComponentNames();

                foreach (var name in names) componentComboBox.Items.Add(name);
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

            componentComboBox.SelectedIndex = 0;
        }

        private void Sheet11_Shutdown(object sender, EventArgs e)
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
            this.componentComboBox.SelectedIndexChanged += new System.EventHandler(this.componentComboBox_SelectedIndexChanged);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.ActivateComponentCombobox);
            this.Shutdown += new System.EventHandler(this.Sheet11_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var thermoSystem = NeqSimThermoSystem.getThermoSystem(); //.clone();
            thermoSystem.setTemperature(thermoSystem.getPhase(0).getComponent(componentComboBox.SelectedItem.ToString())
                .getTriplePointTemperature());
            thermoSystem.setSolidPhaseCheck(false);
            thermoSystem.setSolidPhaseCheck(componentComboBox.SelectedItem.ToString());

            var rangeClear = Range["E2", "L100"];
            rangeClear.Clear();
            var range = Range["A2", "A100"];
            var number = 2;

            int writeStartCell = 1, writeEndCell = 8;

            foreach (Range r in range.Cells)
            {
                var text = (string) r.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    var textVar = "A" + number;
                    var temperatureVar = "B" + number;
                    number++;
                    thermoSystem.setPressure(Range[textVar].Value2);
                    var ops = new ThermodynamicOperations(thermoSystem);
                    ops.freezingPointTemperatureFlash();
                    Range[temperatureVar].Value2 = thermoSystem.getTemperature() - 273.15;
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

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void componentComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*
            componentComboBox.Items.Clear();

            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                var names = thermoSystem.getComponentNames();

                foreach (var name in names) componentComboBox.Items.Add(name);
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

            componentComboBox.SelectedIndex = 0;
            */
        }
    }
}