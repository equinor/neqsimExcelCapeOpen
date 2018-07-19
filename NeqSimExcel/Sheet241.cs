using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using neqsim.thermo.system;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet24
    {
        private readonly List<string> propertyNames = new List<string>();

        private void Sheet24_Startup(object sender, EventArgs e)
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

            propertyNames.ForEach(delegate(string name) { calculationComboBox.Items.Add(name); });


            calculationComboBox.SelectedIndex = 0;

            comp1ComboBox.Items.Clear();


            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                var names = thermoSystem.getComponentNames();

                foreach (var name in names) comp1ComboBox.Items.Add(name);
                ;
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

            comp2ComboBox.Items.Clear();


            thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                var names = thermoSystem.getComponentNames();

                foreach (var name in names)
                    if (!comp1ComboBox.SelectedItem.Equals(name))
                        comp2ComboBox.Items.Add(name);
                ;
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void Sheet24_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            comp2ComboBox.Click += comp2ComboBox_SelectedIndexChanged;
            comp1ComboBox.Click += comp1ComboBox_SelectedIndexChanged;
            calculationComboBox.Click += calculationComboBox_SelectedIndexChanged;
            button1.Click += button1_Click;
            Startup += Sheet24_Startup;
            Shutdown += Sheet24_Shutdown;
        }

        #endregion

        private void calculationComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            calculationComboBox.Items.Clear();


            propertyNames.ForEach(delegate(string name) { calculationComboBox.Items.Add(name); });


            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                var names = thermoSystem.getComponentNames();

                foreach (var name in names)
                    if (!calculationComboBox.Items.Contains("wt fraction " + name))
                        calculationComboBox.Items.Add("wt fraction " + name);
                ;
                foreach (var name in names)
                    if (!calculationComboBox.Items.Contains("activity coefficient " + name))
                        calculationComboBox.Items.Add("activity coefficient " + name);
                ;
                foreach (var name in names)
                    if (!calculationComboBox.Items.Contains("fugacity coefficient " + name))
                        calculationComboBox.Items.Add("fugacity coefficient " + name);
                ;
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void comp1ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            comp1ComboBox.Items.Clear();


            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                var names = thermoSystem.getComponentNames();

                foreach (var name in names) comp1ComboBox.Items.Add(name);
                ;
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void comp2ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            comp2ComboBox.Items.Clear();


            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                var names = thermoSystem.getComponentNames();

                foreach (var name in names)
                    if (!comp1ComboBox.SelectedItem.Equals(name))
                        comp2ComboBox.Items.Add(name);
                ;
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var thermoSystem = (SystemInterface) NeqSimThermoSystem.getThermoSystem().clone();

            var clearRange = (Range) this.get_Range(Cells[8, 1], Cells[1000, 3]);
            clearRange.Clear();

            double temperature = Range["B2"].Value2 + 273.15;
            double pressure = Range["B3"].Value2;
            double NPpoints = Range["B4"].Value2;
            var step = 1.0 / NPpoints;

            thermoSystem.reset();
            var comp1Numb = thermoSystem.getPhase(0).getComponent(comp1ComboBox.SelectedItem.ToString())
                .getComponentNumber();
            var comp2Numb = thermoSystem.getPhase(0).getComponent(comp2ComboBox.SelectedItem.ToString())
                .getComponentNumber();

            thermoSystem.addComponent(comp1Numb, 1.0e-20);
            thermoSystem.addComponent(comp2Numb, 1.0);


            thermoSystem.init(0);

            thermoSystem.setTemperature(temperature);
            thermoSystem.setPressure(pressure);


            for (var j = 0; j < NPpoints + 1; j++)
            {
                var comp1Range = (Range) Cells[8 + j, 1];
                var comp2Range = (Range) Cells[8 + j, 2];
                var setRange = (Range) Cells[8 + j, 3];

                comp1Range.Value2 = step * j;
                comp2Range.Value2 = 1.0 - step * j;

                //   thermoSystem.addComponent(comp1Numb, step);
                //  thermoSystem.addComponent(comp2Numb, -step);

                thermoSystem.init(0);
                var ops = new ThermodynamicOperations(thermoSystem);
                ops.TPflash();

                thermoSystem.init(2);
                thermoSystem.initPhysicalProperties();

                var value = "0";
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
                    value = (thermoSystem.getCp() /
                             (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass() * 1000.0)).ToString();
                }
                else if (calculationComboBox.SelectedItem.Equals("heat capacity Cv"))
                {
                    value = (thermoSystem.getCv() /
                             (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass() * 1000.0)).ToString();
                }
                else if (calculationComboBox.SelectedItem.ToString().Contains("activity coefficient"))
                {
                    var name = calculationComboBox.SelectedItem.ToString().Replace("activity coefficient ", "");
                    value = thermoSystem.getPhase(0)
                        .getActivityCoefficient(thermoSystem.getPhase(0).getComponent(name).getComponentNumber())
                        .ToString();
                    ;
                }
                else if (calculationComboBox.SelectedItem.ToString().Contains("fugacity coefficient"))
                {
                    var name = calculationComboBox.SelectedItem.ToString().Replace("fugacity coefficient ", "");
                    value = thermoSystem.getPhase(0).getComponent(name).getFugasityCoefficient().ToString();
                }
                else if (calculationComboBox.SelectedItem.ToString().Contains("wt fraction"))
                {
                    var name = calculationComboBox.SelectedItem.ToString().Replace("wt fraction ", "");
                    value = (thermoSystem.getPhase(0).getComponent(name).getx() *
                             thermoSystem.getPhase(0).getComponent(name).getMolarMass() /
                             thermoSystem.getPhase(0).getMolarMass()).ToString();
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
                    value = (thermoSystem.getTemperature() - 273.15).ToString();
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