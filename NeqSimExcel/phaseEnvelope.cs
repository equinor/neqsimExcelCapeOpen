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
using Microsoft.Office.Interop.Excel;
namespace NeqSimExcel
{
    public partial class Sheet5
    {

        private void Sheet5_Startup(object sender, System.EventArgs e)
        {
            Excel.Range rangeClear = this.Range["D1", "N300"];
            rangeClear.Clear();
        }

        private void Sheet5_Shutdown(object sender, System.EventArgs e)
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
            this.Startup += new System.EventHandler(this.Sheet5_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet5_Shutdown);
        }

     

        #endregion


        private void button1_Click(object sender, EventArgs e)
        {
            double minPressure = this.Range["B3"].Value2;
            double maxPressure = this.Range["B4"].Value2;

            Excel.Range statusRange = this.Range["B18"];
            
            try
            {
            Excel.Range rangeClear = this.Range["D1", "R300"];
            rangeClear.Clear();
            SystemInterface thermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem().clone();
            thermoSystem.setTemperature(300.0);
            thermoSystem.setPressure(minPressure);

            statusRange.Value2 = "calculating...";
                ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
                ops.setRunAsThread(true);
                Boolean hasWater = thermoSystem.getPhase(0).hasComponent("water");

                double[][] waterData = null;

                if (aqueousCheckBox.Checked && hasWater)
                {
                    ops.waterDewPointLine(minPressure, maxPressure);
                    Boolean isFinished = ops.waitAndCheckForFinishedCalculation(15000);
                    waterData = ops.getData();
                }
                double[][] hydData = null;
                if (hydrateCheckBox.Checked && hasWater)
                {
                    ops.hydrateEquilibriumLine(minPressure, maxPressure);
                    Boolean isFinished = ops.waitAndCheckForFinishedCalculation(15000);
                    hydData = ops.getData();
                }

                if (hasWater)
                {
                    thermoSystem.addComponent("water", -thermoSystem.getPhase(0).getComponent("water").getNumberOfmoles());
                    if(thermoSystem.getPhase(0).hasComponent("MEG")) thermoSystem.addComponent("MEG", -thermoSystem.getPhase(0).getComponent("MEG").getNumberOfmoles());
                    if (thermoSystem.getPhase(0).hasComponent("TEG")) thermoSystem.addComponent("TEG", -thermoSystem.getPhase(0).getComponent("TEG").getNumberOfmoles());
                }

                int number = 1;

                if (aqueousCheckBox.Checked && hasWater)
                {
                    foreach (double val in waterData[0])
                    {
                        string textVar = "O" + number.ToString();
                        string textVar2 = "P" + number.ToString();
                        this.Range[textVar].Value2 = val - 273.15;
                        this.Range[textVar2].Value2 = waterData[1][number - 1];
                        number++;
                    }
                }

                number = 1;
                if (hydrateCheckBox.Checked && hasWater)
                {
                    foreach (double val in hydData[0])
                    {
                        string textVar = "Q" + number.ToString();
                        string textVar2 = "R" + number.ToString();
                        this.Range[textVar].Value2 = val - 273.15;
                        this.Range[textVar2].Value2 = hydData[1][number - 1];
                        number++;
                    }
                }

                var charts = this.ChartObjects() as
                 Microsoft.Office.Interop.Excel.ChartObjects;
                var chartObject = charts.Add(200, 10, 500, 300) as
                 Microsoft.Office.Interop.Excel.ChartObject;
                var chart = chartObject.Chart;

                object misValue = System.Reflection.Missing.Value;
                Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection(misValue);


                if (hydrocarboncheckBox.Checked)
                {
                    try
                    {
                        ops.calcPTphaseEnvelope(true);
                        Boolean isFinished = ops.waitAndCheckForFinishedCalculation(15000);

                    }
                    catch (Exception er)
                    {
                        statusRange.Value2 = "calculation error...";
                        //er.StackTrace.ToString();
                        return;
                    }

                    double[] dewTPoints = ops.get("dewT");
                    double[] dewPPoints = ops.get("dewP");

                    double[] bubTPoints = ops.get("bubT");
                    double[] bubPPoints = ops.get("bubP");

                    double[] criticalPoint = ops.get("criticalPoint1");

                    double[] cricondenbar = ops.get("cricondenbar");
                    double[] cricondenterm = ops.get("cricondentherm");




                    number = 1;
                    foreach (double val in bubTPoints)
                    {
                        string textVar = "D" + number.ToString();
                        string textVar2 = "E" + number.ToString();
                        this.Range[textVar].Value2 = val - 273.15;
                        this.Range[textVar2].Value2 = bubPPoints[number - 1];
                        number++;
                    }

                    number = 1;
                    foreach (double val in dewTPoints)
                    {
                        string textVar = "F" + number.ToString();
                        string textVar2 = "G" + number.ToString();
                        this.Range[textVar].Value2 = val - 273.15;
                        this.Range[textVar2].Value2 = dewPPoints[number - 1];
                        number++;
                    }
                    this.Range["H1"].Value2 = criticalPoint[0] - 273.15;
                    this.Range["I1"].Value2 = criticalPoint[1];

                    this.Range["J1"].Value2 = cricondenbar[0] - 273.15;
                    this.Range["K1"].Value2 = cricondenbar[1];

                    this.Range["L1"].Value2 = cricondenterm[0] - 273.15;
                    this.Range["M1"].Value2 = cricondenterm[1];

                

                   


                    Excel.Series series1 = seriesCollection.NewSeries();
                    series1.XValues = this.Range["D1", "D200"];
                    series1.Values = this.Range["E1", "E200"];
                    series1.Name = "buble point";

                    Excel.Series series2 = seriesCollection.NewSeries();
                    series2.XValues = this.Range["F1", "F200"];
                    series2.Values = this.Range["G1", "G200"];
                    series2.Name = "dew point";

                    Excel.Series series3 = seriesCollection.NewSeries();
                    series3.XValues = this.Range["H1"];
                    series3.Values = this.Range["I1"];
                    series3.Name = "crictical point";

                    Excel.Series series4 = seriesCollection.NewSeries();
                    series4.XValues = this.Range["J1"];
                    series4.Values = this.Range["K1"];
                    series4.Name = "cricondenbar";

                    Excel.Series series5 = seriesCollection.NewSeries();
                    series5.XValues = this.Range["L1"];
                    series5.Values = this.Range["M1"];
                    series5.Name = "cricondentherm";
                }

                if (hydrateCheckBox.Checked && hasWater)
                {
                    Excel.Series seriesHyd = seriesCollection.NewSeries();
                    seriesHyd.XValues = this.Range["Q1", "Q200"];
                    seriesHyd.Values = this.Range["R1", "R200"];
                    seriesHyd.Name = "hydrate equilibrium line";
                }

                if (aqueousCheckBox.Checked && hasWater)
                {
                    Excel.Series seriesAq = seriesCollection.NewSeries();
                    seriesAq.XValues = this.Range["O1", "O200"];
                    seriesAq.Values = this.Range["P1", "P200"];
                    seriesAq.Name = "aqueous dew point line";
                }

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
                    Title: "Phase envelope",
                    CategoryTitle: "Temperature [C]",
                    ValueTitle: "Pressure [bara]");

                statusRange.Value2 = "calculation finished ok.";
            }
            catch (Exception er)
            {
                //er.StackTrace.ToString();
                return;
            }
        }

       


    }
}
