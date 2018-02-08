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
    public partial class Sheet13
    {
        List<string> propertyNames = new List<string>();
        List<string> phaseNames = new List<string>();

        private void Sheet13_Startup(object sender, System.EventArgs e)
        {
           
            phaseNames.Add("overall");
            phaseNames.Add("gas");
            phaseNames.Add("oil");
            phaseNames.Add("aqueous");

            phaseNames.ForEach(delegate(String name)
            {
                phaseNameCombobox.Items.Add(name);
            });
            phaseNameCombobox.SelectedIndex = 0;

            propertyNames.Add("density");
            propertyNames.Add("viscosity");
            propertyNames.Add("compressibility");
            propertyNames.Add("JouleThomson coef.");
            propertyNames.Add("heat capacity Cp");
            propertyNames.Add("heat capacity Cv");
            propertyNames.Add("enthalpy");
            propertyNames.Add("entropy");
            propertyNames.Add("phase fraction (mole)");
            propertyNames.Add("phase fraction (volume)");
            propertyNames.Add("phase fraction (mass)");
            propertyNames.Add("number of phases");
            propertyNames.Add("gas-oil interfacial tension");
            propertyNames.Add("gas-aqueous interfacial tension");
            propertyNames.Add("oil-aqueous interfacial tension");
           
            propertyNames.ForEach(delegate(String name)
            {
                propertyComboBox.Items.Add(name);
            });

           
            propertyComboBox.SelectedIndex = 0;

        }

        private void Sheet13_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.propertyComboBox.SelectedIndexChanged += new System.EventHandler(this.propertyComboBox_SelectedIndexChanged);
            this.propertyComboBox.Click += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            this.phaseNameCombobox.SelectedIndexChanged += new System.EventHandler(this.phaseNameCombobox_SelectedIndexChanged);
            this.okbutton.Click += new System.EventHandler(this.okbutton_Click);
            this.Startup += new System.EventHandler(this.Sheet13_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet13_Shutdown);

        }

        #endregion

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            propertyComboBox.Items.Clear();

          
            propertyNames.ForEach(delegate(String name)
            {
                propertyComboBox.Items.Add(name);
            });


            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                string[] names = thermoSystem.getComponentNames();

                foreach (string name in names)
                {
                    if (!propertyComboBox.Items.Contains("mole fraction " + name)) propertyComboBox.Items.Add("mole fraction " + name);
                }
                foreach (string name in names)
                {
                    if (!propertyComboBox.Items.Contains("wt fraction " + name)) propertyComboBox.Items.Add("wt fraction " + name);
                };
                foreach (string name in names)
                {
                    if (!propertyComboBox.Items.Contains("activity coefficient " + name))  propertyComboBox.Items.Add("activity coefficient " + name);
                };
                foreach (string name in names)
                {
                    if (!propertyComboBox.Items.Contains("fugacity coefficient " + name)) propertyComboBox.Items.Add("fugacity coefficient " + name);
                };
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void okbutton_Click(object sender, EventArgs e)
        {
            SystemInterface thermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem();//.clone();

            Excel.Range clearRange = (Excel.Range)this.get_Range(this.Cells[7, 1], this.Cells[100, 100]);
            clearRange.Clear();

            Double minTemp = this.Range["B2"].Value2;
            Double maxTemp = this.Range["B3"].Value2;
            Double NPTpoints = this.Range["B4"].Value2;
            int numb = 0;
            Excel.Range temprange = this.Range["A8", "A" + (NPTpoints + 7).ToString()];
            foreach (Excel.Range r in temprange.Cells)
            {
                if (NPTpoints > 1)
                    r.Value2 = minTemp + numb * (maxTemp - minTemp) / (NPTpoints - 1);
                else r.Value2 = minTemp;
                numb++;
            }

            Double minPres = this.Range["C2"].Value2;
            Double maxPres = this.Range["C3"].Value2;
            Double NPPpoints = this.Range["C4"].Value2;
             numb = 0;

           // Excel.Range presrange = this.Range["B7", "B" + (NPPpoints + 6).ToString()];
            //foreach (Excel.Range r in presrange.Cells)
                for (int j = 0; j < NPPpoints; j++)
            {
               // r.Value2 = minPres + numb * (maxPres - minPres) / NPPpoints;
                Excel.Range rng = (Excel.Range)this.Cells[7, 2+j];
                if (NPPpoints > 1)
                rng.Value2 = minPres + j * (maxPres - minPres) / (NPPpoints-1);
                else rng.Value2 = minPres;
            }

            for (int i = 0; i < NPTpoints; i++)
            {
                Excel.Range tempera = (Excel.Range)this.Cells[8 + i, 1];
                for (int j = 0; j < NPPpoints; j++)
                {

                    Excel.Range presur = (Excel.Range)this.Cells[7, j+2];
                    thermoSystem.init(0);

                    thermoSystem.checkStability(true);
                    thermoSystem.setTemperature(tempera.Value2+273.15);
                    thermoSystem.setPressure(presur.Value2);
                    ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
                    ops.TPflash();
                    thermoSystem.init(3);
                    thermoSystem.initPhysicalProperties();
                    Excel.Range setRange = (Excel.Range)this.Cells[8 + i, j+2];


                    if (thermoSystem.hasPhaseType(phaseNameCombobox.SelectedItem.ToString()))
                    {
                        string phaseType = phaseNameCombobox.SelectedItem.ToString();
                        string value = "0";
                        if(propertyComboBox.SelectedItem.Equals("density")){
                            value = thermoSystem.getPhaseOfType(phaseType).getPhysicalProperties().getDensity().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("viscosity"))
                        {
                            value = thermoSystem.getPhaseOfType(phaseType).getPhysicalProperties().getViscosity().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("compressibility"))
                        {
                            value = thermoSystem.getPhaseOfType(phaseType).getZ().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("JouleThomson coef."))
                        {
                            value = thermoSystem.getPhaseOfType(phaseType).getJouleThomsonCoefficient().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("heat capacity Cp"))
                        {
                            value = (thermoSystem.getPhaseOfType(phaseType).getCp() / (thermoSystem.getPhaseOfType(phaseType).getMolarMass() * thermoSystem.getPhaseOfType(phaseType).getNumberOfMolesInPhase() * 1000.0)).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("heat capacity Cv"))
                        {
                            value = (thermoSystem.getPhaseOfType(phaseType).getCv() / (thermoSystem.getPhaseOfType(phaseType).getMolarMass() * thermoSystem.getPhaseOfType(phaseType).getNumberOfMolesInPhase() * 1000.0)).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("enthalpy"))
                        {
                            value = (thermoSystem.getPhaseOfType(phaseType).getEnthalpy() / (thermoSystem.getPhaseOfType(phaseType).getMolarMass() * thermoSystem.getPhaseOfType(phaseType).getNumberOfMolesInPhase() * 1000.0)).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("entropy"))
                        {
                            value = (thermoSystem.getPhaseOfType(phaseType).getEntropy() / (thermoSystem.getPhaseOfType(phaseType).getMolarMass() * thermoSystem.getPhaseOfType(phaseType).getNumberOfMolesInPhase() * 1000.0)).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("mole fraction"))
                        {
                            string name = propertyComboBox.SelectedItem.ToString().Replace("mole fraction ", "");
                            value = thermoSystem.getPhaseOfType(phaseType).getComponent(name).getx().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("fugacity coefficient"))
                        {
                            string name = propertyComboBox.SelectedItem.ToString().Replace("fugacity coefficient ", "");
                            value = thermoSystem.getPhaseOfType(phaseType).getComponent(name).getFugasityCoefficient().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("wt fraction"))
                        {
                            string name = propertyComboBox.SelectedItem.ToString().Replace("wt fraction ", "");
                            value = (thermoSystem.getPhaseOfType(phaseType).getComponent(name).getx() * thermoSystem.getPhaseOfType(phaseType).getComponent(name).getMolarMass() / thermoSystem.getPhaseOfType(phaseType).getMolarMass()).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("activity coefficient"))
                        {
                            string name = propertyComboBox.SelectedItem.ToString().Replace("activity coefficient ", "");
                            value = thermoSystem.getPhaseOfType(phaseType).getActivityCoefficient(thermoSystem.getPhaseOfType(phaseType).getComponent(name).getComponentNumber()).ToString(); ;
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("phase fraction (mole)"))
                        {
                            string name = propertyComboBox.SelectedItem.ToString().Replace("phase fraction (mole) ", "");
                            value = thermoSystem.getPhaseOfType(phaseType).getBeta().ToString(); ;
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("phase fraction (volume)"))
                        {
                            string name = propertyComboBox.SelectedItem.ToString().Replace("phase fraction (volume) ", "");
                            double val = thermoSystem.getPhaseOfType(phaseType).getVolume() / thermoSystem.getVolume();
                            value = val.ToString(); ;
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("phase fraction (mass)"))
                        {
                            string name = propertyComboBox.SelectedItem.ToString().Replace("phase fraction (mass) ", "");
                            double val = thermoSystem.getPhaseOfType(phaseType).getNumberOfMolesInPhase() * thermoSystem.getPhaseOfType(phaseType).getMolarMass() / (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass());
                            value = val.ToString(); ;
                        }

                        else value = propertyComboBox.SelectedItem + " not defined for " + phaseNameCombobox.SelectedItem + " phase";
                        setRange.Value2 = value;
                        continue;
                    }

                    if (phaseNameCombobox.SelectedItem.Equals("overall"))
                    {
                        string phaseType = phaseNameCombobox.SelectedItem.ToString();
                        string value = "0";
                        if (propertyComboBox.SelectedItem.Equals("density"))
                        {
                            value = thermoSystem.getDensity().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("compressibility"))
                        {
                            value = thermoSystem.getZ().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("viscosity"))
                        {
                            value = thermoSystem.getViscosity().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("heat capacity Cp"))
                        {
                            value = (thermoSystem.getCp() / (thermoSystem.getTotalNumberOfMoles()*thermoSystem.getMolarMass()*1000.0)).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("heat capacity Cv"))
                        {
                            //value = (thermoSystem.getCv() / (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass() * 1000.0)).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("enthalpy"))
                        {
                            value = (thermoSystem.getEnthalpy() / (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass() * 1000.0)).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("entropy"))
                        {
                            value = (thermoSystem.getEntropy() / (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass()*1000.0)).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("mole fraction"))
                        {
                            string name = propertyComboBox.SelectedItem.ToString().Replace("mole fraction ", "");
                            value = thermoSystem.getPhase(0).getComponent(name).getz().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("number of phases"))
                        {
                            value = thermoSystem.getNumberOfPhases().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Equals("gas-oil interfacial tension") && thermoSystem.hasPhaseType("gas") && thermoSystem.hasPhaseType("oil"))
                        {
                            thermoSystem.calcInterfaceProperties();
                            value = thermoSystem.getInterphaseProperties().getSurfaceTension(thermoSystem.getPhaseNumberOfPhase("gas"), thermoSystem.getPhaseNumberOfPhase("oil")).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Equals("gas-aqueous interfacial tension") && thermoSystem.hasPhaseType("gas") && thermoSystem.hasPhaseType("aqueous"))
                        {
                            thermoSystem.calcInterfaceProperties();
                            value = thermoSystem.getInterphaseProperties().getSurfaceTension(thermoSystem.getPhaseNumberOfPhase("gas"), thermoSystem.getPhaseNumberOfPhase("aqueous")).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Equals("oil-aqueous interfacial tension") && thermoSystem.hasPhaseType("oil") && thermoSystem.hasPhaseType("aqueous"))
                        {
                            thermoSystem.calcInterfaceProperties();
                            value = thermoSystem.getInterphaseProperties().getSurfaceTension(thermoSystem.getPhaseNumberOfPhase("oil"), thermoSystem.getPhaseNumberOfPhase("aqueous")).ToString();
                        }
                        else value = propertyComboBox.SelectedItem + " not defined for " + phaseNameCombobox.SelectedItem + " phase";
                        setRange.Value2 = value;
                        continue;
                    }
                    setRange.Value2 = "no " + phaseNameCombobox.SelectedItem + " phase";

                }
            }

          
           

         //   int number = 1;
          //  foreach (Excel.Range r in range.Cells)
          //  {
           //     string text = (string)r.Text;

              //  thermoSystem.setTemperature(r.Value2 + 273.15);
               // thermoSystem.setPressure(this.Range[textVar].Value2);
               // ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
               // ops.TPflash();
            
        }

        private void propertyComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void phaseNameCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
