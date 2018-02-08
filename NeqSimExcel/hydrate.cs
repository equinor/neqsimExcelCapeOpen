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
    public partial class hydrate
    {
        private void Sheet4_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet4_Shutdown(object sender, System.EventArgs e)
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
            this.Startup += new System.EventHandler(this.Sheet4_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet4_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Range rangeClear = this.Range["D2", "N300"];
            rangeClear.Clear();


            Excel.Range rangeMessage = this.Range["C7"];
            rangeMessage.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            rangeMessage.Value2 = "calculating....please wait...";
            Cursor previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            if (!thermoSystem.getPhase(0).hasComponent("water"))
            {
                rangeMessage.Value2 = "no water in fluid....water must be present in hydrate calculations...";
                return;
            }
            thermoSystem.setTemperature(273.75);
            thermoSystem.setHydrateCheck(true);
            Excel.Range range = this.Range["A2", "A100"];
            int number = 1;
            int writeStartCell = 1, writeEndCell = 8;
            
            foreach (Excel.Range r in range.Cells)
            {
                string text = (string)r.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    number++;
                    thermoSystem.setPressure(r.Value2);
                    ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
                    ops.hydrateFormationTemperature();
                    string textVar = "B" + number.ToString();
                    this.Range[textVar].Value2 = thermoSystem.getTemperature() - 273.15;

                    var table = thermoSystem.createTable("fluid");
                    int rows = table.Length;
                    int columns = table[1].Length;
                    writeEndCell = writeStartCell + rows;

                    var startCell = Cells[writeStartCell, 7];
                    var endCell = Cells[writeEndCell - 1, columns + 6];
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

            rangeMessage.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
            rangeMessage.Value2 = "calculations finished..ok";
            Cursor.Current = previousCursor;


            var charts = this.ChartObjects() as
             Microsoft.Office.Interop.Excel.ChartObjects;
            var chartObject = charts.Add(10, 200, 500, 400) as
                Microsoft.Office.Interop.Excel.ChartObject;
            var chart = chartObject.Chart;

            object misValue = System.Reflection.Missing.Value;

            Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection(misValue);

            Excel.Series series1 = seriesCollection.NewSeries();
            series1.XValues = this.Range["B2", "B200"];
            series1.Values = this.Range["A2", "A200"];
            series1.Name = "hydrate equilibrium line";

            Excel.Axis axisY = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue,
        Excel.XlAxisGroup.xlPrimary);
            axisY.HasTitle = true;
            axisY.AxisTitle.Text = "Pressure [bara]";

            Excel.Axis axisX = (Excel.Axis)chartObject.Chart.Axes(Excel.XlAxisType.xlCategory,
        Excel.XlAxisGroup.xlPrimary);
            axisX.HasTitle = true;
            axisX.AxisTitle.Text = "Temperature [°C]";

            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterSmooth;
            chart.ChartWizard(Source: seriesCollection, HasLegend: false,
                Title: "Hydrate equilibrium temperatures",
                CategoryTitle: "Temperature [C]",
                ValueTitle: "Pressure [bara]");
            

        }
    }
}
