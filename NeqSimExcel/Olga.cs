using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Olga
    {
        private void Sheet7_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet7_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            this.pathLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.pathLabel_LinkClicked);
            this.Startup += new System.EventHandler(this.Sheet7_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet7_Shutdown);

        }

        #endregion


        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
        }

        private void pathLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (fileOLGAdialog.ShowDialog() == DialogResult.OK) Range["B8"].Value2 = fileOLGAdialog.FileName;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            var rangeMessage = Range["B14"];
            rangeMessage.Font.Color = ColorTranslator.ToOle(Color.Blue);
            rangeMessage.Value2 = "calculating OLGA table....please wait...";
            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            var ops = new ThermodynamicOperations(thermoSystem);

            double minTemp = Range["B4"].Value2 + 273.15;
            double maxTemp = Range["B5"].Value2 + 273.15;
            int tempSteps = Convert.ToInt32(Range["B6"].Value2);

            double minPres = Range["C4"].Value2;
            double maxPres = Range["C5"].Value2;
            int pressteps = Convert.ToInt32(Range["C6"].Value2);

            string filename = Range["B8"].Value2;

            ops.OLGApropTable(minTemp, maxTemp, tempSteps, minPres, maxPres, pressteps, filename, 0);
            rangeMessage.Font.Color = ColorTranslator.ToOle(Color.Green);
            try
            {
                ops.displayResult();
                rangeMessage.Value2 = "calculating OLGA table finished ok";
                Process.Start(filename);
            }
            catch (Exception ex)
            {
                ex.ToString();
                rangeMessage.Value2 = "error creating file " + filename;
            }
        }
    }
}