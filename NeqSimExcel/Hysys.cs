using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Principal;
using System.Windows.Forms;
using DatabaseConnection;
using DatabaseConnection.NeqSimDatabaseSetTableAdapters;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet6
    {
        //  string[] packageNames = new string[]{""};

        private void Sheet6_Startup(object sender, EventArgs e)
        {
            try
            {
                var test = new fluidinfoTableAdapter();

                //NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter test = new NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter test = new NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfo1TableAdapter test = new neqsimdatabaseDataSetTableAdapters.fluidinfo1TableAdapter();

                var userName = WindowsIdentity.GetCurrent().Name;
                userName = userName.Replace("STATOIL-NET\\", "");
                userName = userName.Replace("WIN-NTNU-NO\\", "");
                userName = userName.ToLower();


                //         NeqSimExcel.DataSet1.fluidinfoDataTable tt = test.GetData(userName);
                var tt = test.GetDataBy(userName);

                var names = new List<string>();
                names.Add("New fluid");
                fluidListNameComboBox.Items.Add("New fluid");
                //names.Add("CPApackage");
                //names.Add(WindowsIdentity.GetCurrent().Name);
                foreach (NeqSimDatabaseSet.fluidinfoRow row in tt.Rows)
                {
                    var tempString = "";
                    try
                    {
                        tempString = row.TEXT;
                    }
                    catch (Exception exept)
                    {
                        tempString = "";
                        exept.ToString();
                    }
                    finally
                    {
                    }

                    names.Add(row.ID + " " + tempString);
                    fluidListNameComboBox.Items.Add(row.ID + " " + tempString);
                }


                //   packageNames = names.ToArray();
                //   fluidListNameComboBox.Items.Add(names.ToList());
                fluidListNameComboBox.SelectedIndex = 0;
            }
            catch (Exception excet)
            {
                Console.WriteLine("Error " + excet.Message);
            }

            sharedCheckbox.Visible = true; // hide if another fluid is chosen!
        }


    private void Sheet6_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.fluidListNameComboBox.MouseClick += new System.Windows.Forms.MouseEventHandler(this.fluidListNameComboBox_MouseClick);
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            this.Startup += new System.EventHandler(this.Sheet6_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet6_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var rangeClear = Range["E2", "L100"];
            rangeClear.Clear();

            var textVar1 = "B9";
            Range[textVar1].Value2 = "saving fluid...";
            int a2;
            var tab = new fluidinfoTableAdapter();
            var usertab = new userdbTableAdapter();

            if (fluidListNameComboBox.SelectedItem.ToString().Equals("New fluid"))
            {
                sharedCheckbox.Visible = true;
                a2 = Convert.ToInt32(tab.InsertNewFluidRow());
                var userName = WindowsIdentity.GetCurrent().Name;
                userName = userName.Replace("STATOIL-NET\\", "");
                userName = userName.Replace("WIN-NTNU-NO\\", "");
                userName = userName.ToLower();
                var userid = Convert.ToInt32(usertab.getUserID(userName));
                tab.UpdateUserID(userid, a2);
                tab.UpdateField("new field", a2);
                tab.UpdateWell("new well", a2);
                tab.UpdateTest("", a2);
                tab.UpdateSample("", a2);
                tab.UpdateHistory("", a2);
                if (sharedCheckbox.Checked) tab.UpdateShared(1, a2);
            }
            else
            {
                a2 = Convert.ToInt32(fluidListNameComboBox.SelectedItem.ToString().Split(' ')[0]);
            }

            // Adding a description of the fluid
            if (Range["B5"].Value2 != null)
            {
                string description = Range["B5"].Value2;
                //   DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter tab = new DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter();
                tab.UpdateDescription(description, a2);
            }

            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            thermoSystem.saveFluid(a2);

            Range[textVar1].Value2 = "finished saved fluidID " + a2;

            int writeStartCell = 1, writeEndCell = 8;


            thermoSystem.init(0);
            thermoSystem.setTemperature(288.15);
            thermoSystem.setPressure(1.01325);
            var ops = new ThermodynamicOperations(thermoSystem);
            ops.TPflash();

            var table = thermoSystem.createTable("fluid");
            var rows = table.Length;
            var columns = table[1].Length;
            writeEndCell = writeStartCell + rows;

            var startCell = Cells[writeStartCell, 6];
            var endCell = Cells[writeEndCell - 1, columns + 5];
            var writeRange = Range[startCell, endCell];

            writeStartCell += rows + 3;
            //writeRange.Value2 = table;

            var data = new object[rows, columns];
            for (var row = 1; row <= rows; row++)
            for (var column = 1; column <= columns; column++)
                data[row - 1, column - 1] = table[row - 1][column - 1];

            writeRange.Value2 = data;
        }


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
                string path = Range["B18"].Value2;

                string filePath = null;

                var textVar1 = "B25";
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

                string fluidName = Range["B20"].Value2 + ".neqsim";

                // string fullname = path + "/"+fluidName;
                var fullname = filePath + "/" + fluidName;

                var thermoSystem = NeqSimThermoSystem.getThermoSystem();
                thermoSystem.saveObjectToFile(fullname, "");

                Range["B25"].Value2 = "Saved fluid.." + Range["B20"].Value2;
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
            fluidListNameComboBox.Items.Clear();
            sharedCheckbox.Visible = true;

            var test = new fluidinfoTableAdapter();
            //            NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter test = new NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter();

            var userName = WindowsIdentity.GetCurrent().Name;
            userName = userName.Replace("STATOIL-NET\\", "");
            userName = userName.Replace("WIN-NTNU-NO\\", "");
            userName = userName.ToLower();
            var tt = test.GetDataBy(userName);
            // NeqSimExcel.DataSet1.fluidinfoDataTable tt = test.GetData(userName);
            var names = new List<string>();
            names.Add("New fluid");
            fluidListNameComboBox.Items.Add("New fluid");
            //names.Add("CPApackage");
            //names.Add(WindowsIdentity.GetCurrent().Name);
            foreach (NeqSimDatabaseSet.fluidinfoRow row in tt.Rows)
            {
                var tempString = "";
                try
                {
                    tempString = row.TEXT;
                }
                catch (Exception exept)
                {
                    tempString = "";
                    exept.ToString();
                }
                finally
                {
                }

                names.Add(row.ID + " " + tempString);
                fluidListNameComboBox.Items.Add(row.ID + " " + tempString);
            }

        }
    }
}