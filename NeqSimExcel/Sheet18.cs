﻿using System;
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

namespace NeqSimExcel
{
    public partial class Sheet18
    {
        private void Sheet18_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet18_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.gasViscosityCommboBox.SelectedIndexChanged += new System.EventHandler(this.EoScombobox_SelectedIndexChanged);
            this.oilViscosityCmboBox.SelectedIndexChanged += new System.EventHandler(this.oilViscosityCmboBox_SelectedIndexChanged);
            this.aqueousViscositycomboBox.SelectedIndexChanged += new System.EventHandler(this.aqueousViscositycomboBox_SelectedIndexChanged);
            this.comboBox3.SelectedIndexChanged += new System.EventHandler(this.comboBox3_SelectedIndexChanged);
            this.comboBox4.SelectedIndexChanged += new System.EventHandler(this.comboBox4_SelectedIndexChanged);
            this.comboBox5.SelectedIndexChanged += new System.EventHandler(this.comboBox5_SelectedIndexChanged);
            this.oilCOnductivityComboBox.SelectedIndexChanged += new System.EventHandler(this.oilCOnductivityComboBox_SelectedIndexChanged);
            this.gasCondutivityComboBox.SelectedIndexChanged += new System.EventHandler(this.comboBox8_SelectedIndexChanged);
            this.aqueousConduvtivityComboBox.SelectedIndexChanged += new System.EventHandler(this.aqueousConduvtivityComboBox_SelectedIndexChanged);
            this.Startup += new System.EventHandler(this.Sheet18_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet18_Shutdown);

        }

        #endregion

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeInterfacialTensionGasOil();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeInterfacialTensionGasAqueous();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeInterfacialTensionOilAqueous();
        }

        public void changeInterfacialTensionGasOil()
        {
            if (comboBox3.SelectedItem.ToString().Equals("Linear Gradient Theory"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("gas", "oil", "Linear Gradient Theory");
            }
            else if (comboBox3.SelectedItem.ToString().Equals("Parachor"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("gas", "oil", "Parachor");
            }
            else if (comboBox3.SelectedItem.ToString().Equals("Firozabadi Ramley"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("gas", "oil", "Firozabadi Ramley");
            }
            else if (comboBox3.SelectedItem.ToString().Equals("Full Gradient Theory"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("gas", "oil", "Full Gradient Theory");
            }
            else if (comboBox3.SelectedItem.ToString().Equals("Simple Gradient Theory"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("gas", "oil", "Simple Gradient Theory");
            }
        }

        public void changeInterfacialTensionGasAqueous()
        {
            if (comboBox4.SelectedItem.ToString().Equals("Linear Gradient Theory"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("gas", "aqueous", "Linear Gradient Theory");
            }
            else if (comboBox4.SelectedItem.ToString().Equals("Parachor"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("gas", "aqueous", "Parachor");
            }
            else if (comboBox4.SelectedItem.ToString().Equals("Firozabadi Ramley"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("gas", "aqueous", "Firozabadi Ramley");
            }
            else if (comboBox4.SelectedItem.ToString().Equals("Full Gradient Theory"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("gas", "aqueous", "Full Gradient Theory");
            }
            else if (comboBox4.SelectedItem.ToString().Equals("Simple Gradient Theory"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("gas", "aqueous", "Simple Gradient Theory");
            }
        }

        public void changeInterfacialTensionOilAqueous()
        {
            if (comboBox5.SelectedItem.ToString().Equals("Linear Gradient Theory"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("oil", "aqueous", "Linear Gradient Theory");
            }
            else if (comboBox5.SelectedItem.ToString().Equals("Parachor"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("oil", "aqueous", "Parachor");
            }
            else if (comboBox5.SelectedItem.ToString().Equals("Firozabadi Ramley"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("oil", "aqueous", "Firozabadi Ramley");
            }
            else if (comboBox5.SelectedItem.ToString().Equals("Full Gradient Theory"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("oil", "aqueous", "Full Gradient Theory");
            }
            else if (comboBox5.SelectedItem.ToString().Equals("Simple Gradient Theory"))
            {
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties().setInterfacialTensionModel("oil", "aqueous", "Simple Gradient Theory");
            }
        }



        private void EoScombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
          
            changeViscosityModel("gas", gasViscosityCommboBox.SelectedItem.ToString());
        }

        public void changeViscosityModel(string phase, string method)
        {
            if (NeqSimThermoSystem.getThermoSystem().hasPhaseType(phase))
            {
                NeqSimThermoSystem.getThermoSystem().getPhaseOfType(phase).getPhysicalProperties().setViscosityModel(method);
            }
        }

        private void oilViscosityCmboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeViscosityModel("oil", oilViscosityCmboBox.SelectedItem.ToString());
        }

        private void aqueousViscositycomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeViscosityModel("aqueous", aqueousViscositycomboBox.SelectedItem.ToString());
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeConductivityModel("gas", gasCondutivityComboBox.SelectedItem.ToString());
        }

        public void changeConductivityModel(string phase, string method)
        {
            if (NeqSimThermoSystem.getThermoSystem().hasPhaseType(phase))
            {
                NeqSimThermoSystem.getThermoSystem().getPhaseOfType(phase).getPhysicalProperties().setConductivityModel(method);
            }
        }

        private void oilCOnductivityComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeConductivityModel("oil", oilCOnductivityComboBox.SelectedItem.ToString());
        }

        private void aqueousConduvtivityComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeConductivityModel("aqueous", aqueousConduvtivityComboBox.SelectedItem.ToString());
        }
    }
}
