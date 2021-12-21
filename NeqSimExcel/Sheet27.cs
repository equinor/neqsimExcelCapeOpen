using System;
using System.Collections.Generic;
using System.Security.Principal;
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
          
        }

        private void ActivateWorkSheet()
        {
            selectLocalFLuidCOmboBox.Items.Clear();
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
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.ActivateWorkSheet);
            this.Startup += new System.EventHandler(this.Sheet27_Startup);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var textVar1 = "B8";
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

        }

        private void selectFluidCombobox_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {

        }
    }
}