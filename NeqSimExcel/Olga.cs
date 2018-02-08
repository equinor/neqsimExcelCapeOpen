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
    public partial class Olga
    {
        private void Sheet7_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet7_Shutdown(object sender, System.EventArgs e)
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
            this.fileOLGAdialog.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            this.pathLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.pathLabel_LinkClicked);
            this.Startup += new System.EventHandler(this.Sheet7_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet7_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {

         
            Excel.Range rangeMessage = this.Range["B14"];
            rangeMessage.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            rangeMessage.Value2 = "calculating OLGA table....please wait...";
            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);

            double  minTemp = this.Range["B4"].Value2 + 273.15;
            double maxTemp = this.Range["B5"].Value2 + 273.15;
            int tempSteps = Convert.ToInt32(this.Range["B6"].Value2);

            double minPres = this.Range["C4"].Value2;
            double maxPres = this.Range["C5"].Value2;
            int pressteps = Convert.ToInt32(this.Range["C6"].Value2);

            string filename = this.Range["B8"].Value2;

            ops.OLGApropTable(minTemp, maxTemp, tempSteps, minPres, maxPres, pressteps, filename, 0);
            rangeMessage.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
            try
            {
                ops.displayResult();
                rangeMessage.Value2 = "calculating OLGA table finished ok";
                System.Diagnostics.Process.Start(filename);
            }
            catch (Exception ex)
            {
                ex.ToString();
                rangeMessage.Value2 = "error creating file " + filename;
            }
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void pathLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (fileOLGAdialog.ShowDialog() == DialogResult.OK)
            {
                this.Range["B8"].Value2 = this.fileOLGAdialog.FileName;
            }
        }

    }
}
