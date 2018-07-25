using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class hydrate
    {
        private void Sheet4_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet4_Shutdown(object sender, EventArgs e)
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
            this.Startup += new System.EventHandler(this.Sheet4_Startup);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var rangeClear = Range["D2", "N300"];
            rangeClear.Clear();


            var rangeMessage = Range["C7"];
            rangeMessage.Font.Color = ColorTranslator.ToOle(Color.Blue);
            rangeMessage.Value2 = "calculating....please wait...";
            var previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            if (!thermoSystem.getPhase(0).hasComponent("water"))
            {
                rangeMessage.Value2 = "no water in fluid....water must be present in hydrate calculations...";
                return;
            }

            thermoSystem.setTemperature(273.75);
            thermoSystem.setHydrateCheck(true);
            var range = Range["A2", "A100"];
            var number = 1;
            int writeStartCell = 1, writeEndCell = 8;

            foreach (Range r in range.Cells)
            {
                var text = (string) r.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    number++;
                    thermoSystem.setPressure(r.Value2);
                    var ops = new ThermodynamicOperations(thermoSystem);
                    ops.hydrateFormationTemperature();
                    var textVar = "B" + number;
                    Range[textVar].Value2 = thermoSystem.getTemperature() - 273.15;

                    var table = thermoSystem.createTable("fluid");
                    var rows = table.Length;
                    var columns = table[1].Length;
                    writeEndCell = writeStartCell + rows;

                    var startCell = Cells[writeStartCell, 7];
                    var endCell = Cells[writeEndCell - 1, columns + 6];
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

            rangeMessage.Font.Color = ColorTranslator.ToOle(Color.Green);
            rangeMessage.Value2 = "calculations finished..ok";
            Cursor.Current = previousCursor;


            var charts = ChartObjects() as
                ChartObjects;
            var chartObject = charts.Add(10, 200, 500, 400);
            var chart = chartObject.Chart;

            object misValue = Missing.Value;

            var seriesCollection = (SeriesCollection) chart.SeriesCollection(misValue);

            var series1 = seriesCollection.NewSeries();
            series1.XValues = Range["B2", "B200"];
            series1.Values = Range["A2", "A200"];
            series1.Name = "hydrate equilibrium line";

            var axisY = (Axis) chart.Axes(XlAxisType.xlValue,
                XlAxisGroup.xlPrimary);
            axisY.HasTitle = true;
            axisY.AxisTitle.Text = "Pressure [bara]";

            var axisX = (Axis) chartObject.Chart.Axes(XlAxisType.xlCategory,
                XlAxisGroup.xlPrimary);
            axisX.HasTitle = true;
            axisX.AxisTitle.Text = "Temperature [°C]";

            chart.ChartType = XlChartType.xlXYScatterSmooth;
            chart.ChartWizard(seriesCollection, HasLegend: false,
                Title: "Hydrate equilibrium temperatures",
                CategoryTitle: "Temperature [C]",
                ValueTitle: "Pressure [bara]");
        }
    }
}