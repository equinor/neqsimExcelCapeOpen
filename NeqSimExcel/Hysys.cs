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
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Collections;
using System.Security.Principal;
using MySql.Data;
using DatabaseConnection;

namespace NeqSimExcel
{
    public partial class Sheet6   {
       //  string[] packageNames = new string[]{""};

        private void Sheet6_Startup(object sender, System.EventArgs e)
        {
            try
            {
                DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter test = new DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter();
          
                //NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter test = new NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter test = new NeqSimExcelDataSetTableAdapters.fluidinfoTableAdapter();
                //NeqSimExcelDataSetTableAdapters.fluidinfo1TableAdapter test = new neqsimdatabaseDataSetTableAdapters.fluidinfo1TableAdapter();

                string userName = WindowsIdentity.GetCurrent().Name;
                userName = userName.Replace("STATOIL-NET\\", "");
                userName = userName.Replace("WIN-NTNU-NO\\", "");
                userName = userName.ToLower();


       //         NeqSimExcel.DataSet1.fluidinfoDataTable tt = test.GetData(userName);
                DatabaseConnection.NeqSimDatabaseSet.fluidinfoDataTable tt = test.GetDataBy(userName);
       
                List<string> names = new List<string>();
                names.Add("New fluid");
                fluidListNameComboBox.Items.Add("New fluid");
                //names.Add("CPApackage");
                //names.Add(WindowsIdentity.GetCurrent().Name);
                foreach (DatabaseConnection.NeqSimDatabaseSet.fluidinfoRow row in tt.Rows)
                {
                    string tempString = "";
                    try
                    {
                        tempString = row.TEXT.ToString();
                    }
                    catch (Exception exept)
                    {
                        tempString = "";
                        exept.ToString();
                    }
                    finally
                    {


                    }
                    names.Add(row.ID.ToString() + " " + tempString);
                    fluidListNameComboBox.Items.Add(row.ID.ToString() + " " + tempString);

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

        private void Sheet6_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.fluidListNameComboBox.SelectedIndexChanged += new System.EventHandler(this.fluidListNameComboBox_SelectedIndexChanged_1);
            this.fluidListNameComboBox.Click += new System.EventHandler(this.fluidListNameComboBox_SelectedIndexChanged);
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            this.Startup += new System.EventHandler(this.Sheet6_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet6_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Range rangeClear = this.Range["E2", "L100"];
            rangeClear.Clear();

            string textVar1 = "B9";
            this.Range[textVar1].Value2 = "saving fluid...";
            int a2;
            DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter tab = new DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter();
            DatabaseConnection.NeqSimDatabaseSetTableAdapters.userdbTableAdapter usertab = new DatabaseConnection.NeqSimDatabaseSetTableAdapters.userdbTableAdapter();

            if (fluidListNameComboBox.SelectedItem.ToString().Equals("New fluid")){

                sharedCheckbox.Visible = true;
                a2 = Convert.ToInt32((object)tab.InsertNewFluidRow());
                string userName = WindowsIdentity.GetCurrent().Name;
                userName = userName.Replace("STATOIL-NET\\", "");
                userName = userName.Replace("WIN-NTNU-NO\\", "");
                userName = userName.ToLower();
                int userid = Convert.ToInt32((object)usertab.getUserID(userName));
                tab.UpdateUserID(userid, a2);
                tab.UpdateField("new field", a2);
                tab.UpdateWell("new well", a2);
                tab.UpdateTest("", a2);
                tab.UpdateSample("", a2);
                tab.UpdateHistory("", a2);
                if (sharedCheckbox.Checked == true)
                {
                    tab.UpdateShared(1, a2);
                }

            }
            else{
                a2 = Convert.ToInt32((String)fluidListNameComboBox.SelectedItem.ToString().Split(' ')[0]);
            }

            // Adding a description of the fluid
            if (this.Range["D3"].Value2 != null)
            { 
                string description = this.Range["D3"].Value2;
             //   DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter tab = new DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter();
                tab.UpdateDescription(description, a2);
            }
 
            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            thermoSystem.saveFluid(a2);

             this.Range[textVar1].Value2 = "finished saved fluidID " + a2.ToString();
            
             int writeStartCell = 1, writeEndCell = 8;
            

                     thermoSystem.init(0);
                     thermoSystem.setTemperature(288.15);
                     thermoSystem.setPressure(1.01325);
                     ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
                     ops.TPflash();

                     var table = thermoSystem.createTable("fluid");
                     int rows = table.Length;
                     int columns = table[1].Length;
                     writeEndCell = writeStartCell + rows;

                     var startCell = Cells[writeStartCell, 6];
                     var endCell = Cells[writeEndCell - 1, columns + 5];
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


        private void fluidListNameComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
        
            fluidListNameComboBox.Items.Clear();
            sharedCheckbox.Visible = true;

            DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter test = new DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter();
//            NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter test = new NeqSimExcel.DataSet1TableAdapters.fluidinfoTableAdapter();
           
            string userName = WindowsIdentity.GetCurrent().Name;
            userName = userName.Replace("STATOIL-NET\\", "");
            userName = userName.Replace("WIN-NTNU-NO\\", "");
            userName = userName.ToLower();
            DatabaseConnection.NeqSimDatabaseSet.fluidinfoDataTable tt = test.GetDataBy(userName);
           // NeqSimExcel.DataSet1.fluidinfoDataTable tt = test.GetData(userName);
            List<string> names = new List<string>();
            names.Add("New fluid");
            fluidListNameComboBox.Items.Add("New fluid");
            //names.Add("CPApackage");
            //names.Add(WindowsIdentity.GetCurrent().Name);
            foreach (DatabaseConnection.NeqSimDatabaseSet.fluidinfoRow row in tt.Rows)
            {
                string tempString = "";
                try
                {
                    tempString = row.TEXT.ToString();
                }
                catch (Exception exept)
                {
                    tempString = "";
                    exept.ToString();
                }
                finally
                {


                }
                names.Add(row.ID.ToString() + " " + tempString);
                fluidListNameComboBox.Items.Add(row.ID.ToString() + " " + tempString);

            }
            //   packageNames = names.ToArray();
            //   fluidListNameComboBox.Items.Add(names.ToList());

        }


        private void fluidListNameComboBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog openFileDIalog = new FolderBrowserDialog();
            openFileDIalog.ShowDialog();
            string localFileName = openFileDIalog.SelectedPath;
            this.Range["B18"].Value2 = localFileName;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            string path = this.Range["B18"].Value2;

            string filePath = null;

            if (NeqSimThermoSystem.LocalFilePath == null)
            {
                filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                filePath = filePath + "/AppData/Roaming/neqsim/fluids/";
            }
            else
            {
                filePath = NeqSimThermoSystem.LocalFilePath;
            }


            string fluidName = this.Range["B20"].Value2 + ".neqsim";

           // string fullname = path + "/"+fluidName;
            string fullname = filePath + "/" + fluidName;

            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            thermoSystem.saveObjectToFile(fullname,"");

            this.Range["B25"].Value2 = "Saved fluid.."+ this.Range["B20"].Value2;
        }

    }
}
