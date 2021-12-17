using System;
using System.Reflection;
using ikvm.extensions;
using Microsoft.Office.Interop.Excel;
using neqsim.thermo.system;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet5
    {
        private void Sheet5_Startup(object sender, EventArgs e)
        {
          
        }

        private void Sheet5_Shutdown(object sender, EventArgs e)
        {
        }

         private new void Activate()
        {
            var rangeClear = Range["D1", "N300"];
            rangeClear.Clear();
        }

        #region VSTO Designer generated code

            /// <summary>
            ///     Required method for Designer support - do not modify
            ///     the contents of this method with the code editor.
            /// </summary>
            private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.Activate);
            this.Startup += new System.EventHandler(this.Sheet5_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet5_Shutdown);

        }

        #endregion


        private void button1_Click(object sender, EventArgs e)
        {
            double minPressure = Range["B3"].Value2;
            double maxPressure = Range["B4"].Value2;

            var statusRange = Range["B18"];

            try
            {
                var rangeClear = Range["D1", "R300"];
                rangeClear.Clear();
                var thermoSystem = (SystemInterface) NeqSimThermoSystem.getThermoSystem().clone();
                thermoSystem.setTemperature(300.0);
                thermoSystem.setPressure(minPressure);

                statusRange.Value2 = "calculating...";
                var ops = new ThermodynamicOperations(thermoSystem);
                ops.setRunAsThread(true);
                var hasWater = thermoSystem.getPhase(0).hasComponent("water");

                double[][] waterData = null;

                if (aqueousCheckBox.Checked && hasWater)
                {
                    ops.waterDewPointLine(minPressure, maxPressure);
                    var isFinished = ops.waitAndCheckForFinishedCalculation(15000);
                    waterData = ops.getData();
                }

                double[][] hydData = null;
                if (hydrateCheckBox.Checked && hasWater)
                {
                    ops.hydrateEquilibriumLine(minPressure, maxPressure);
                    var isFinished = ops.waitAndCheckForFinishedCalculation(15000);
                    hydData = ops.getData();
                }

                if (hasWater)
                {
                    thermoSystem.addComponent("water",
                        -thermoSystem.getPhase(0).getComponent("water").getNumberOfmoles());
                    if (thermoSystem.getPhase(0).hasComponent("MEG"))
                        thermoSystem.addComponent("MEG",
                            -thermoSystem.getPhase(0).getComponent("MEG").getNumberOfmoles());
                    if (thermoSystem.getPhase(0).hasComponent("TEG"))
                        thermoSystem.addComponent("TEG",
                            -thermoSystem.getPhase(0).getComponent("TEG").getNumberOfmoles());
                }

                var number = 1;

                if (aqueousCheckBox.Checked && hasWater)
                    foreach (var val in waterData[0])
                    {
                        var textVar = "O" + number;
                        var textVar2 = "P" + number;
                        Range[textVar].Value2 = val - 273.15;
                        Range[textVar2].Value2 = waterData[1][number - 1];
                        number++;
                    }

                number = 1;
                if (hydrateCheckBox.Checked && hasWater)
                    foreach (var val in hydData[0])
                    {
                        var textVar = "Q" + number;
                        var textVar2 = "R" + number;
                        Range[textVar].Value2 = val - 273.15;
                        Range[textVar2].Value2 = hydData[1][number - 1];
                        number++;
                    }

                var charts = ChartObjects() as
                    ChartObjects;
                var chartObject = charts.Add(200, 10, 500, 300);
                var chart = chartObject.Chart;

                object misValue = Missing.Value;
                var seriesCollection = (SeriesCollection) chart.SeriesCollection(misValue);


                if (hydrocarboncheckBox.Checked)
                {
                    try
                    {
                        ops.calcPTphaseEnvelope();
                        var isFinished = ops.waitAndCheckForFinishedCalculation(15000);
                    }
                    catch (Exception er)
                    {
                        statusRange.Value2 = "calculation error..." + er.Message;
                        //er.StackTrace.ToString();
                        return;
                    }

                    var dewTPoints = ops.get("dewT");
                    var dewPPoints = ops.get("dewP");

                    var dewTPoints2 = new double[0];
                    var dewPPoints2 = new double[0];
                    try
                    {
                        dewTPoints2 = ops.get("dewT2");
                        dewPPoints2 = ops.get("dewP2");
                    }
                    catch(Exception er)
                    {
                        er.printStackTrace();
                    }

                    var bubTPoints = ops.get("bubT");
                    var bubPPoints = ops.get("bubP");

                    var bubTPoints2 = new double[0];
                    var bubPPoints2 = new double[0];
                    try
                    {
                        bubTPoints2 = ops.get("bubT2");
                        bubPPoints2 = ops.get("bubP2");
                    }
                    catch(Exception err)
                    {
                        err.printStackTrace();
                    }

                    var criticalPoint = ops.get("criticalPoint1");

                    var cricondenbar = ops.get("cricondenbar");
                    var cricondenterm = ops.get("cricondentherm");


                    number = 1;
                    foreach (var val in bubTPoints)
                    {
                        var textVar = "D" + number;
                        var textVar2 = "E" + number;
                        Range[textVar].Value2 = val - 273.15;
                        Range[textVar2].Value2 = bubPPoints[number - 1];
                        number++;
                    }

                    int number2 = 1;
                    foreach (var val in bubTPoints2)
                    {
                        var textVar = "D" + number;
                        var textVar2 = "E" + number;
                        Range[textVar].Value2 = val - 273.15;
                        Range[textVar2].Value2 = bubPPoints2[number2 - 1];
                        number++;
                        number2++;
                    }

                    number = 1;
                    foreach (var val in dewTPoints)
                    {
                        var textVar = "F" + number;
                        var textVar2 = "G" + number;
                        Range[textVar].Value2 = val - 273.15;
                        Range[textVar2].Value2 = dewPPoints[number - 1];
                        number++;
                    }

                    number2 = 1;
                    foreach (var val in dewTPoints2)
                    {
                        var textVar = "F" + number;
                        var textVar2 = "G" + number;
                        Range[textVar].Value2 = val - 273.15;
                        Range[textVar2].Value2 = dewPPoints2[number2 - 1];
                        number++;
                        number2++;
                    }

                    Range["H1"].Value2 = criticalPoint[0] - 273.15;
                    Range["I1"].Value2 = criticalPoint[1];

                    Range["J1"].Value2 = cricondenbar[0] - 273.15;
                    Range["K1"].Value2 = cricondenbar[1];

                    Range["L1"].Value2 = cricondenterm[0] - 273.15;
                    Range["M1"].Value2 = cricondenterm[1];


                    var series1 = seriesCollection.NewSeries();
                    series1.XValues = Range["D1", "D200"];
                    series1.Values = Range["E1", "E200"];
                    series1.Name = "buble point";

                    var series2 = seriesCollection.NewSeries();
                    series2.XValues = Range["F1", "F200"];
                    series2.Values = Range["G1", "G200"];
                    series2.Name = "dew point";

                    var series3 = seriesCollection.NewSeries();
                    series3.XValues = Range["H1"];
                    series3.Values = Range["I1"];
                    series3.Name = "crictical point";

                    var series4 = seriesCollection.NewSeries();
                    series4.XValues = Range["J1"];
                    series4.Values = Range["K1"];
                    series4.Name = "cricondenbar";

                    var series5 = seriesCollection.NewSeries();
                    series5.XValues = Range["L1"];
                    series5.Values = Range["M1"];
                    series5.Name = "cricondentherm";
                }

                if (hydrateCheckBox.Checked && hasWater)
                {
                    var seriesHyd = seriesCollection.NewSeries();
                    seriesHyd.XValues = Range["Q1", "Q200"];
                    seriesHyd.Values = Range["R1", "R200"];
                    seriesHyd.Name = "hydrate equilibrium line";
                }

                if (aqueousCheckBox.Checked && hasWater)
                {
                    var seriesAq = seriesCollection.NewSeries();
                    seriesAq.XValues = Range["O1", "O200"];
                    seriesAq.Values = Range["P1", "P200"];
                    seriesAq.Name = "aqueous dew point line";
                }

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
                    Title: "Phase envelope",
                    CategoryTitle: "Temperature [C]",
                    ValueTitle: "Pressure [bara]");

                statusRange.Value2 = "calculation finished ok.";
            }
            catch (Exception er)
            {
                statusRange.Value2 = "calculation error..." + er.Message;
                //er.StackTrace.ToString();
            }
        }
    }
}