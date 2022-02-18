using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using neqsim.thermodynamicOperations;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet13
    {
        private readonly List<string> phaseNames = new List<string>();
        private readonly List<string> propertyNames = new List<string>();

        private void Sheet13_Startup(object sender, EventArgs e)
        {
            phaseNames.Add("overall");
            phaseNames.Add("gas");
            phaseNames.Add("oil");
            phaseNames.Add("aqueous");

            phaseNames.ForEach(delegate(string name) { phaseNameCombobox.Items.Add(name); });
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

            propertyNames.ForEach(delegate(string name) { propertyComboBox.Items.Add(name); });


            propertyComboBox.SelectedIndex = 0;
        }

        private void Sheet13_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.propertyComboBox.Click += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            this.okbutton.Click += new System.EventHandler(this.okbutton_Click);
            this.Startup += new System.EventHandler(this.Sheet13_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet13_Shutdown);

        }

        #endregion

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            propertyComboBox.Items.Clear();


            propertyNames.ForEach(delegate(string name) { propertyComboBox.Items.Add(name); });


            var thermoSystem = NeqSimThermoSystem.getThermoSystem();
            try
            {
                var names = thermoSystem.getComponentNames();

                foreach (var name in names)
                    if (!propertyComboBox.Items.Contains("mole fraction " + name))
                        propertyComboBox.Items.Add("mole fraction " + name);
                foreach (var name in names)
                    if (!propertyComboBox.Items.Contains("wt fraction " + name))
                        propertyComboBox.Items.Add("wt fraction " + name);
                ;
                foreach (var name in names)
                    if (!propertyComboBox.Items.Contains("activity coefficient " + name))
                        propertyComboBox.Items.Add("activity coefficient " + name);
                ;
                foreach (var name in names)
                    if (!propertyComboBox.Items.Contains("fugacity coefficient " + name))
                        propertyComboBox.Items.Add("fugacity coefficient " + name);
                ;
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void okbutton_Click(object sender, EventArgs e)
        {
            var thermoSystem = NeqSimThermoSystem.getThermoSystem(); //.clone();

            var clearRange = (Range) this.get_Range(Cells[7, 1], Cells[100, 100]);
            clearRange.Clear();

            double minTemp = Range["B2"].Value2;
            double maxTemp = Range["B3"].Value2;
            double NPTpoints = Range["B4"].Value2;
            var numb = 0;
            var temprange = Range["A8", "A" + (NPTpoints + 7)];
            foreach (Range r in temprange.Cells)
            {
                if (NPTpoints > 1)
                    r.Value2 = minTemp + numb * (maxTemp - minTemp) / (NPTpoints - 1);
                else r.Value2 = minTemp;
                numb++;
            }

            double minPres = Range["C2"].Value2;
            double maxPres = Range["C3"].Value2;
            double NPPpoints = Range["C4"].Value2;
            numb = 0;

            // Excel.Range presrange = this.Range["B7", "B" + (NPPpoints + 6).ToString()];
            //foreach (Excel.Range r in presrange.Cells)
            for (var j = 0; j < NPPpoints; j++)
            {
                // r.Value2 = minPres + numb * (maxPres - minPres) / NPPpoints;
                var rng = (Range) Cells[7, 2 + j];
                if (NPPpoints > 1)
                    rng.Value2 = minPres + j * (maxPres - minPres) / (NPPpoints - 1);
                else rng.Value2 = minPres;
            }

            for (var i = 0; i < NPTpoints; i++)
            {
                var tempera = (Range) Cells[8 + i, 1];
                for (var j = 0; j < NPPpoints; j++)
                {
                    var presur = (Range) Cells[7, j + 2];
                    thermoSystem.init(0);

                    thermoSystem.checkStability(true);
                    thermoSystem.setTemperature(tempera.Value2 + 273.15);
                    thermoSystem.setPressure(presur.Value2);
                    var ops = new ThermodynamicOperations(thermoSystem);
                    ops.TPflash();
                    thermoSystem.init(3);
                    thermoSystem.initPhysicalProperties();
                    var setRange = (Range) Cells[8 + i, j + 2];


                    if (thermoSystem.hasPhaseType(phaseNameCombobox.SelectedItem.ToString()))
                    {
                        var phaseType = phaseNameCombobox.SelectedItem.ToString();
                        var value = "0";
                        if (propertyComboBox.SelectedItem.Equals("density"))
                        {
                            value = thermoSystem.getPhaseOfType(phaseType).getPhysicalProperties().getDensity()
                                .ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("viscosity"))
                        {
                            value = thermoSystem.getPhaseOfType(phaseType).getPhysicalProperties().getViscosity()
                                .ToString();
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
                            value = (thermoSystem.getPhaseOfType(phaseType).getCp() /
                                     (thermoSystem.getPhaseOfType(phaseType).getMolarMass() *
                                      thermoSystem.getPhaseOfType(phaseType).getNumberOfMolesInPhase() * 1000.0))
                                .ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("heat capacity Cv"))
                        {
                            value = (thermoSystem.getPhaseOfType(phaseType).getCv() /
                                     (thermoSystem.getPhaseOfType(phaseType).getMolarMass() *
                                      thermoSystem.getPhaseOfType(phaseType).getNumberOfMolesInPhase() * 1000.0))
                                .ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("enthalpy"))
                        {
                            value = (thermoSystem.getPhaseOfType(phaseType).getEnthalpy() /
                                     (thermoSystem.getPhaseOfType(phaseType).getMolarMass() *
                                      thermoSystem.getPhaseOfType(phaseType).getNumberOfMolesInPhase() * 1000.0))
                                .ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("entropy"))
                        {
                            value = (thermoSystem.getPhaseOfType(phaseType).getEntropy() /
                                     (thermoSystem.getPhaseOfType(phaseType).getMolarMass() *
                                      thermoSystem.getPhaseOfType(phaseType).getNumberOfMolesInPhase() * 1000.0))
                                .ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("mole fraction"))
                        {
                            var name = propertyComboBox.SelectedItem.ToString().Replace("mole fraction ", "");
                            value = thermoSystem.getPhaseOfType(phaseType).getComponent(name).getx().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("fugacity coefficient"))
                        {
                            var name = propertyComboBox.SelectedItem.ToString().Replace("fugacity coefficient ", "");
                            value = thermoSystem.getPhaseOfType(phaseType).getComponent(name).getFugacityCoefficient()
                                .ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("wt fraction"))
                        {
                            var name = propertyComboBox.SelectedItem.ToString().Replace("wt fraction ", "");
                            value = (thermoSystem.getPhaseOfType(phaseType).getComponent(name).getx() *
                                     thermoSystem.getPhaseOfType(phaseType).getComponent(name).getMolarMass() /
                                     thermoSystem.getPhaseOfType(phaseType).getMolarMass()).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("activity coefficient"))
                        {
                            var name = propertyComboBox.SelectedItem.ToString().Replace("activity coefficient ", "");
                            value = thermoSystem.getPhaseOfType(phaseType)
                                .getActivityCoefficient(thermoSystem.getPhaseOfType(phaseType).getComponent(name)
                                    .getComponentNumber()).ToString();
                            ;
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("phase fraction (mole)"))
                        {
                            var name = propertyComboBox.SelectedItem.ToString().Replace("phase fraction (mole) ", "");
                            value = thermoSystem.getPhaseOfType(phaseType).getBeta().ToString();
                            ;
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("phase fraction (volume)"))
                        {
                            var name = propertyComboBox.SelectedItem.ToString().Replace("phase fraction (volume) ", "");
                            var val = thermoSystem.getPhaseOfType(phaseType).getVolume() / thermoSystem.getVolume();
                            value = val.ToString();
                            ;
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("phase fraction (mass)"))
                        {
                            var name = propertyComboBox.SelectedItem.ToString().Replace("phase fraction (mass) ", "");
                            var val = thermoSystem.getPhaseOfType(phaseType).getNumberOfMolesInPhase() *
                                      thermoSystem.getPhaseOfType(phaseType).getMolarMass() /
                                      (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass());
                            value = val.ToString();
                            ;
                        }

                        else
                        {
                            value = propertyComboBox.SelectedItem + " not defined for " +
                                    phaseNameCombobox.SelectedItem + " phase";
                        }

                        setRange.Value2 = value;
                        continue;
                    }

                    if (phaseNameCombobox.SelectedItem.Equals("overall"))
                    {
                        var phaseType = phaseNameCombobox.SelectedItem.ToString();
                        var value = "0";
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
                            value = (thermoSystem.getCp() /
                                     (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass() * 1000.0))
                                .ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("heat capacity Cv"))
                        {
                            //value = (thermoSystem.getCv() / (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass() * 1000.0)).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("enthalpy"))
                        {
                            value = (thermoSystem.getEnthalpy() /
                                     (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass() * 1000.0))
                                .ToString();
                        }
                        else if (propertyComboBox.SelectedItem.Equals("entropy"))
                        {
                            value = (thermoSystem.getEntropy() /
                                     (thermoSystem.getTotalNumberOfMoles() * thermoSystem.getMolarMass() * 1000.0))
                                .ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("mole fraction"))
                        {
                            var name = propertyComboBox.SelectedItem.ToString().Replace("mole fraction ", "");
                            value = thermoSystem.getPhase(0).getComponent(name).getz().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Contains("number of phases"))
                        {
                            value = thermoSystem.getNumberOfPhases().ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Equals("gas-oil interfacial tension") &&
                                 thermoSystem.hasPhaseType("gas") && thermoSystem.hasPhaseType("oil"))
                        {
                            thermoSystem.calcInterfaceProperties();
                            value = thermoSystem.getInterphaseProperties()
                                .getSurfaceTension(thermoSystem.getPhaseNumberOfPhase("gas"),
                                    thermoSystem.getPhaseNumberOfPhase("oil")).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Equals("gas-aqueous interfacial tension") &&
                                 thermoSystem.hasPhaseType("gas") && thermoSystem.hasPhaseType("aqueous"))
                        {
                            thermoSystem.calcInterfaceProperties();
                            value = thermoSystem.getInterphaseProperties()
                                .getSurfaceTension(thermoSystem.getPhaseNumberOfPhase("gas"),
                                    thermoSystem.getPhaseNumberOfPhase("aqueous")).ToString();
                        }
                        else if (propertyComboBox.SelectedItem.ToString().Equals("oil-aqueous interfacial tension") &&
                                 thermoSystem.hasPhaseType("oil") && thermoSystem.hasPhaseType("aqueous"))
                        {
                            thermoSystem.calcInterfaceProperties();
                            value = thermoSystem.getInterphaseProperties()
                                .getSurfaceTension(thermoSystem.getPhaseNumberOfPhase("oil"),
                                    thermoSystem.getPhaseNumberOfPhase("aqueous")).ToString();
                        }
                        else
                        {
                            value = propertyComboBox.SelectedItem + " not defined for " +
                                    phaseNameCombobox.SelectedItem + " phase";
                        }

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

    }
}