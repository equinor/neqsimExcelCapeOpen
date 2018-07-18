﻿using System;
using System.Collections.Generic;
using System.Security.Principal;
using DatabaseConnection;
using DatabaseConnection.NeqSimDatabaseSetTableAdapters;
using thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet27
    {
        private void Sheet27_Startup(object sender, EventArgs e)
        {
            try
            {
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
            calcButton.Click += calcButton_Click;
            Startup += Sheet27_Startup;
            Shutdown += Sheet27_Shutdown;
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
    }
}