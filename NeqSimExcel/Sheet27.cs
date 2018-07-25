using System;
using System.Collections.Generic;
using System.Security.Principal;
using DatabaseConnection;
using DatabaseConnection.NeqSimDatabaseSetTableAdapters;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using Microsoft.Win32;

namespace NeqSimExcel
{
    public partial class Sheet27
    {
        private void Sheet27_Startup(object sender, EventArgs e)
        {
            
            try
            {
                var names = new List<string>();
                var filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                var fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";

                var d = new DirectoryInfo(fullPath);
                var Files = d.GetFiles("*.neqsim");
                foreach (var file in Files)
                {
                    names.Add(file.Name.Replace(".neqsim", ""));
                    selectLocalFLuidCOmboBox.Items.Add(file.Name.Replace(".neqsim", ""));
                }

                selectLocalFLuidCOmboBox.SelectedIndex = 0;

            }
            catch (Exception excet)
            {
                Console.WriteLine("Error " + excet.Message);
            }
            

            try
            {
                var test = new fluidinfoTableAdapter();

                var userName = WindowsIdentity.GetCurrent().Name;
                userName = userName.Replace("STATOIL-NET\\", "");
                //userName = userName.Replace("WIN-NTNU-NO\\", "");
                userName = userName.ToLower();

                var names = new List<string>();
                var tt = test.GetDataBy(userName);
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
                    selectFluidCombobox.Items.Add(row.ID + " " + tempString);
                }

                //   packageNames = names.ToArray();
                //   fluidListNameComboBox.Items.Add(names.ToList());
                selectFluidCombobox.SelectedIndex = 0;
            }
            catch (Exception excet)
            {
                Console.WriteLine("Error " + excet.Message);
            }

         


        }

        private void Sheet27_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.selectFluidCombobox.MouseClick += new System.Windows.Forms.MouseEventHandler(this.selectFluidCombobox_MouseClick);
            this.calcButton.Click += new System.EventHandler(this.calcButton_Click);
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.selectLocalFLuidCOmboBox.Click += new System.EventHandler(this.selectLocalCHanged);
            this.Startup += new System.EventHandler(this.Sheet27_Startup);

        }

        #endregion

        private void calcButton_Click(object sender, EventArgs e)
        {
            var rangeClear = Range["E2", "L100"];
            rangeClear.Clear();

            var textVar1 = "B9";
            Range[textVar1].Value2 = "open fluid...";

            var a2 = Convert.ToInt32(selectFluidCombobox.SelectedItem.ToString().Split(' ')[0]);
            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            thermoSystem = thermoSystem.readObject(a2);

            Range[textVar1].Value2 = "finished reading fluidID " + a2;

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
            NeqSimThermoSystem.setThermoSystem(thermoSystem);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var textVar1 = "B16";
            Range[textVar1].Value2 = "open fluid...";
            string name = selectLocalFLuidCOmboBox.SelectedItem.ToString();
            var filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";
            var filename = fullPath + "\\" + name + ".neqsim";
            NeqSimThermoSystem.setThermoSystem(NeqSimThermoSystem.getThermoSystem().readObjectFromFile(filename, filename));

            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            
            Range[textVar1].Value2 = "finished reading local fluid " + name;

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
            NeqSimThermoSystem.setThermoSystem(thermoSystem);
        }

        private void selectLocalCHanged(object sender, EventArgs e)
        {
            try
            {
                selectLocalFLuidCOmboBox.Items.Clear();
                var names = new List<string>();
                var filePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                var fullPath = filePath + "\\AppData\\Roaming\\neqsim\\fluids";

                var d = new DirectoryInfo(fullPath);
                var Files = d.GetFiles("*.neqsim");
                foreach (var file in Files)
                {
                    names.Add(file.Name.Replace(".neqsim", ""));
                    selectLocalFLuidCOmboBox.Items.Add(file.Name.Replace(".neqsim", ""));
                }
                
                selectLocalFLuidCOmboBox.SelectedIndex = 0;
            }
            catch (Exception excet)
            {
                Console.WriteLine("Error " + excet.Message);
            }

        }

        private void selectFluidCombobox_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            try
            {
                selectFluidCombobox.Items.Clear();
                var test = new fluidinfoTableAdapter();

                var userName = WindowsIdentity.GetCurrent().Name;
                userName = userName.Replace("STATOIL-NET\\", "");
                userName = userName.Replace("WIN-NTNU-NO\\", "");
                userName = userName.ToLower();

                var names = new List<string>();
                var tt = test.GetDataBy(userName);
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

                        names.Add(row.ID + " " + tempString);
                        selectFluidCombobox.Items.Add(row.ID + " " + tempString);
                    }

                }

                //   packageNames = names.ToArray();
                //   fluidListNameComboBox.Items.Add(names.ToList());
                selectFluidCombobox.SelectedIndex = 0;
            }
            catch (Exception excet)
            {
                Console.WriteLine("Error " + excet.Message);
            }

        }
    }
}