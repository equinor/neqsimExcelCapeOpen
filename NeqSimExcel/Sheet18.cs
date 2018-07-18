using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet18
    {
        private void Sheet18_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet18_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            gasViscosityCommboBox.SelectedIndexChanged += EoScombobox_SelectedIndexChanged;
            oilViscosityCmboBox.SelectedIndexChanged += oilViscosityCmboBox_SelectedIndexChanged;
            aqueousViscositycomboBox.SelectedIndexChanged += aqueousViscositycomboBox_SelectedIndexChanged;
            comboBox3.SelectedIndexChanged += comboBox3_SelectedIndexChanged;
            comboBox4.SelectedIndexChanged += comboBox4_SelectedIndexChanged;
            comboBox5.SelectedIndexChanged += comboBox5_SelectedIndexChanged;
            oilCOnductivityComboBox.SelectedIndexChanged += oilCOnductivityComboBox_SelectedIndexChanged;
            gasCondutivityComboBox.SelectedIndexChanged += comboBox8_SelectedIndexChanged;
            aqueousConduvtivityComboBox.SelectedIndexChanged += aqueousConduvtivityComboBox_SelectedIndexChanged;
            button1.Click += button1_Click;
            Startup += Sheet18_Startup;
            Shutdown += Sheet18_Shutdown;
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
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("gas", "oil", "Linear Gradient Theory");
            else if (comboBox3.SelectedItem.ToString().Equals("Parachor"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("gas", "oil", "Parachor");
            else if (comboBox3.SelectedItem.ToString().Equals("Firozabadi Ramley"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("gas", "oil", "Firozabadi Ramley");
            else if (comboBox3.SelectedItem.ToString().Equals("Full Gradient Theory"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("gas", "oil", "Full Gradient Theory");
            else if (comboBox3.SelectedItem.ToString().Equals("Simple Gradient Theory"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("gas", "oil", "Simple Gradient Theory");
        }

        public void changeInterfacialTensionGasAqueous()
        {
            if (comboBox4.SelectedItem.ToString().Equals("Linear Gradient Theory"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("gas", "aqueous", "Linear Gradient Theory");
            else if (comboBox4.SelectedItem.ToString().Equals("Parachor"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("gas", "aqueous", "Parachor");
            else if (comboBox4.SelectedItem.ToString().Equals("Firozabadi Ramley"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("gas", "aqueous", "Firozabadi Ramley");
            else if (comboBox4.SelectedItem.ToString().Equals("Full Gradient Theory"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("gas", "aqueous", "Full Gradient Theory");
            else if (comboBox4.SelectedItem.ToString().Equals("Simple Gradient Theory"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("gas", "aqueous", "Simple Gradient Theory");
        }

        public void changeInterfacialTensionOilAqueous()
        {
            if (comboBox5.SelectedItem.ToString().Equals("Linear Gradient Theory"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("oil", "aqueous", "Linear Gradient Theory");
            else if (comboBox5.SelectedItem.ToString().Equals("Parachor"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("oil", "aqueous", "Parachor");
            else if (comboBox5.SelectedItem.ToString().Equals("Firozabadi Ramley"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("oil", "aqueous", "Firozabadi Ramley");
            else if (comboBox5.SelectedItem.ToString().Equals("Full Gradient Theory"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("oil", "aqueous", "Full Gradient Theory");
            else if (comboBox5.SelectedItem.ToString().Equals("Simple Gradient Theory"))
                NeqSimThermoSystem.getThermoSystem().getInterphaseProperties()
                    .setInterfacialTensionModel("oil", "aqueous", "Simple Gradient Theory");
        }


        private void EoScombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeViscosityModel("gas", gasViscosityCommboBox.SelectedItem.ToString());
        }

        public void changeViscosityModel(string phase, string method)
        {
            if (NeqSimThermoSystem.getThermoSystem().hasPhaseType(phase))
                NeqSimThermoSystem.getThermoSystem().getPhaseOfType(phase).getPhysicalProperties()
                    .setViscosityModel(method);
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
                NeqSimThermoSystem.getThermoSystem().getPhaseOfType(phase).getPhysicalProperties()
                    .setConductivityModel(method);
        }

        private void oilCOnductivityComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeConductivityModel("oil", oilCOnductivityComboBox.SelectedItem.ToString());
        }

        private void aqueousConduvtivityComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            changeConductivityModel("aqueous", aqueousConduvtivityComboBox.SelectedItem.ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var openFileDIalog = new FolderBrowserDialog();
            openFileDIalog.ShowDialog();
            var localFileName = openFileDIalog.SelectedPath;
            Range["B23"].Value2 = localFileName;
            NeqSimThermoSystem.LocalFilePath = localFileName;
        }
    }
}