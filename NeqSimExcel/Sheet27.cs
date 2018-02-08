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

namespace NeqSimExcel
{
    public partial class Sheet27
    {
        private void Sheet27_Startup(object sender, System.EventArgs e)
        {

            try
            {
                DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter test = new DatabaseConnection.NeqSimDatabaseSetTableAdapters.fluidinfoTableAdapter();

                string userName = WindowsIdentity.GetCurrent().Name;
                userName = userName.Replace("STATOIL-NET\\", "");
                userName = userName.Replace("WIN-NTNU-NO\\", "");
                userName = userName.ToLower();

                List<string> names = new List<string>();
                DatabaseConnection.NeqSimDatabaseSet.fluidinfoDataTable tt = test.GetDataBy(userName);
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
                selectFluidCombobox.Items.Add(row.ID.ToString() + " " + tempString);

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

        private void Sheet27_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.calcButton.Click += new System.EventHandler(this.calcButton_Click);
            this.Startup += new System.EventHandler(this.Sheet27_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet27_Shutdown);

        }

        #endregion

        private void calcButton_Click(object sender, EventArgs e)
        {
            Excel.Range rangeClear = this.Range["E2", "L100"];
            rangeClear.Clear();

            string textVar1 = "B9";
            this.Range[textVar1].Value2 = "open fluid...";

            int a2 = Convert.ToInt32((String)selectFluidCombobox.SelectedItem.ToString().Split(' ')[0]);
            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            thermoSystem = thermoSystem.readObject(a2);

            this.Range[textVar1].Value2 = "finished reading fluidID " + a2.ToString();

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

            var startCell = Cells[writeStartCell, 5];
            var endCell = Cells[writeEndCell - 1, columns + 4];
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
            NeqSimThermoSystem.setThermoSystem(thermoSystem);

        }
    }
}
