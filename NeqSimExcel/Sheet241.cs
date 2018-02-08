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
    public partial class Sheet24
    {

        List<string> propertyNames = new List<string>();

        private void Sheet24_Startup(object sender, System.EventArgs e)
        {
            propertyNames.Add("density");
            propertyNames.Add("viscosity");
            propertyNames.Add("compressibility");
            propertyNames.Add("JouleThomson coef.");
            propertyNames.Add("heat capacity Cp");
            propertyNames.Add("heat capacity Cv");
            propertyNames.Add("buble point pressure");
            propertyNames.Add("buble point temperature");
            propertyNames.Add("dew point pressure");
            propertyNames.Add("dew point temperature");

            propertyNames.ForEach(delegate (String name)
            {
                calculationComboBox.Items.Add(name);
            });


            calculationComboBox.SelectedIndex = 0;

            comp1ComboBox.Items.Clear();


            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                string[] names = thermoSystem.getComponentNames();

                foreach (string name in names)
                {
                    comp1ComboBox.Items.Add(name);
                };
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

            comp2ComboBox.Items.Clear();


            thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                string[] names = thermoSystem.getComponentNames();

                foreach (string name in names)
                {
                    if (!comp1ComboBox.SelectedItem.Equals(name)) comp2ComboBox.Items.Add(name);
                };
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

        }

        private void Sheet24_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.comp2ComboBox.Click += new System.EventHandler(this.comp2ComboBox_SelectedIndexChanged);
            this.comp1ComboBox.Click += new System.EventHandler(this.comp1ComboBox_SelectedIndexChanged);
            this.calculationComboBox.Click += new System.EventHandler(this.calculationComboBox_SelectedIndexChanged);
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.Startup += new System.EventHandler(this.Sheet24_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet24_Shutdown);

        }

        #endregion

        private void calculationComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            calculationComboBox.Items.Clear();


            propertyNames.ForEach(delegate (String name)
            {
                calculationComboBox.Items.Add(name);
            });


            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                string[] names = thermoSystem.getComponentNames();

                foreach (string name in names)
                {
                    if (!calculationComboBox.Items.Contains("wt fraction " + name)) calculationComboBox.Items.Add("wt fraction " + name);
                };
                foreach (string name in names)
                {
                    if (!calculationComboBox.Items.Contains("activity coefficient " + name)) calculationComboBox.Items.Add("activity coefficient " + name);
                };
                foreach (string name in names)
                {
                    if (!calculationComboBox.Items.Contains("fugacity coefficient " + name)) calculationComboBox.Items.Add("fugacity coefficient " + name);
                };
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void comp1ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            comp1ComboBox.Items.Clear();
            

            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                string[] names = thermoSystem.getComponentNames();

                foreach (string name in names)
                {
                    comp1ComboBox.Items.Add(name);
                };
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void comp2ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            comp2ComboBox.Items.Clear();


            SystemInterface thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                string[] names = thermoSystem.getComponentNames();

                foreach (string name in names)
                {
                    if(!comp1ComboBox.SelectedItem.Equals(name)) comp2ComboBox.Items.Add(name);
                };
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SystemInterface thermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem().clone();

            Excel.Range clearRange = (Excel.Range)this.get_Range(this.Cells[8, 1], this.Cells[1000, 3]);
            clearRange.Clear();

            Double temperature = this.Range["B2"].Value2+273.15;
            Double pressure = this.Range["B3"].Value2;
            Double NPpoints = this.Range["B4"].Value2;
            double step = 1.0 / NPpoints;

            thermoSystem.reset();
            int comp1Numb = thermoSystem.getPhase(0).getComponent(comp1ComboBox.SelectedItem.ToString()).getComponentNumber();
            int comp2Numb = thermoSystem.getPhase(0).getComponent(comp2ComboBox.SelectedItem.ToString()).getComponentNumber();

            thermoSystem.addComponent(comp1Numb, 1.0e-20);
            thermoSystem.addComponent(comp2Numb, 1.0);
            

            thermoSystem.init(0);

            thermoSystem.setTemperature(temperature);
            thermoSystem.setPressure(pressure);


            for (int j = 0; j < NPpoints+1; j++)
            {
                Excel.Range comp1Range = (Excel.Range)this.Cells[8 + j, 1];
                Excel.Range comp2Range = (Excel.Range)this.Cells[8 + j, 2];
                Excel.Range setRange = (Excel.Range)this.Cells[8 + j, 3];

                comp1Range.Value2 = step*(j);
                comp2Range.Value2 = 1.0-step*(j);

             //   thermoSystem.addComponent(comp1Numb, step);
              //  thermoSystem.addComponent(comp2Numb, -step);

                thermoSystem.init(0);
                ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
                ops.TPflash();

                thermoSystem.init(2);
                thermoSystem.initPhysicalProperties();

                string value = "0";
                if (calculationComboBox.SelectedItem.Equals("density"))
                {
                    value = thermoSystem.getPhase(0).getPhysicalProperties().getDensity().ToString();
                }
                else if (calculationComboBox.SelectedItem.Equals("viscosity"))
                {
                    value = thermoSystem.getPhase(0).getPhysicalProperties().getViscosity().ToString();
                }
                else if (calculationComboBox.SelectedItem.Equals("heat capacity Cp"))
                {
                    value = (thermoSystem.getCp() / (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass() * 1000.0)).ToString();
                }
                else if (calculationComboBox.SelectedItem.Equals("heat capacity Cv"))
                {
                    value = (thermoSystem.getCv() / (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass() * 1000.0)).ToString();
                }
                else if (calculationComboBox.SelectedItem.ToString().Contains("activity coefficient"))
                {
                    string name = calculationComboBox.SelectedItem.ToString().Replace("activity coefficient ", "");
                    value = thermoSystem.getPhase(0).getActivityCoefficient(thermoSystem.getPhase(0).getComponent(name).getComponentNumber()).ToString(); ;
                }
                else if (calculationComboBox.SelectedItem.ToString().Contains("fugacity coefficient"))
                {
                    string name = calculationComboBox.SelectedItem.ToString().Replace("fugacity coefficient ", "");
                    value = thermoSystem.getPhase(0).getComponent(name).getFugasityCoefficient().ToString();
                }
                else if (calculationComboBox.SelectedItem.ToString().Contains("wt fraction"))
                {
                    string name = calculationComboBox.SelectedItem.ToString().Replace("wt fraction ", "");
                    value = (thermoSystem.getPhase(0).getComponent(name).getx() * thermoSystem.getPhase(0).getComponent(name).getMolarMass() / thermoSystem.getPhase(0).getMolarMass()).ToString();
                }
                else if (calculationComboBox.SelectedItem.Equals("JouleThomson coef."))
                {
                    value = thermoSystem.getPhase(0).getJouleThomsonCoefficient().ToString();
                }
                else if (calculationComboBox.SelectedItem.Equals("buble point pressure"))
                {
                    ops.bubblePointPressureFlash(false);
                    value = thermoSystem.getPressure().ToString();
                }
                else if (calculationComboBox.SelectedItem.Equals("buble point temperature"))
                {
                    ops.bubblePointTemperatureFlash();
                    value = (thermoSystem.getTemperature()-273.15).ToString();
                }
                else if (calculationComboBox.SelectedItem.Equals("dew point pressure"))
                {
                    ops.dewPointPressureFlash();
                    value = thermoSystem.getPressure().ToString();
                }
                else if (calculationComboBox.SelectedItem.Equals("dew point temperature"))
                {
                    ops.dewPointTemperatureFlash();
                    value = (thermoSystem.getTemperature() - 273.15).ToString();
                }
                setRange.Value2 = value;


                thermoSystem.addComponent(comp1Numb, step);
                thermoSystem.addComponent(comp2Numb, -step);
            }

            }
    }
}
