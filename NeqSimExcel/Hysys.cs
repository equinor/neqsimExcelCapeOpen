using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Principal;
using System.Windows.Forms;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

using System.Diagnostics;
using System.ComponentModel;

namespace NeqSimExcel
{
    public partial class Sheet6
    {
        //  string[] packageNames = new string[]{""};

        private void Sheet6_Startup(object sender, EventArgs e)
        {
          
        }


    private void Sheet6_Shutdown(object sender, EventArgs e)
        {
        }

        private void ActivateWorksheet()
        {
        }

        #region VSTO Designer generated code

            /// <summary>
            ///     Required method for Designer support - do not modify
            ///     the contents of this method with the code editor.
            /// </summary>
            private void InternalStartup()
        {
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.ActivateWorksheet);
            this.Startup += new System.EventHandler(this.Sheet6_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet6_Shutdown);

        }

        #endregion


        private void fluidListNameComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void fluidListNameComboBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var openFileDIalog = new FolderBrowserDialog();
            openFileDIalog.ShowDialog();
            var localFileName = openFileDIalog.SelectedPath;
            Range["B18"].Value2 = localFileName;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                string filePath = null;

                var textVar1 = "B8";
                var nameUrl = "B3";
                Range[textVar1].Value2 = "saving fluid...";

                if (NeqSimThermoSystem.LocalFilePath == null)
                {
                    filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                    filePath = filePath + "/AppData/Roaming/neqsim/fluids/";
                }
                else
                {
                    filePath = NeqSimThermoSystem.LocalFilePath;
                }

                if (!Directory.Exists(filePath))
                {
                    DirectoryInfo di = Directory.CreateDirectory(filePath);
                }

                string fluidName = Range[nameUrl].Value2 + ".neqsim";

                // string fullname = path + "/"+fluidName;
                var fullname = filePath + "/" + fluidName;

                var thermoSystem = NeqSimThermoSystem.getThermoSystem();
                thermoSystem.saveObjectToFile(fullname, "");

                Range[textVar1].Value2 = "Saved fluid.." + Range[nameUrl].Value2;
            }
            catch (Exception exept)
            {
                Console.WriteLine("The process failed: {0}", exept.ToString());
            }
            finally
            {
            }
        }

        private void fluidListNameComboBox_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";

            Process.Start("explorer.exe",fullPath);
        }
    }
}