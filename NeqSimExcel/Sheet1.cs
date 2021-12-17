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
using neqsim.thermo.system;
using neqsim.thermodynamicOperations;
using neqsim.PVTsimulation.simulation;
using neqsim.PVTsimulation.util.parameterfitting;
using neqsim.statistics.parameterFitting;
using neqsim.statistics.parameterFitting.nonLinearParameterFitting;

namespace NeqSimExcel
{
    public partial class Sheet1
    {
        string plusMWcell = "";
        Boolean forceUseCPAmodel = false;
        Excel.Range statusRange;
        Excel.Range sumMolPrecentRange;

        Excel.Range molPrecentCompRange;
        Excel.Range allComponentRowsRange, extendedComponentRowsRange, extendedComponentRowsRange2, extendedComponentRowsRange3, extendedComponentRowsRange4, extendedComponentRowsRange5, extendedComponentRowsRange6;
        Excel.Range averageMolWtRange;
        Excel.Range TBPCompRange;
        Excel.Range waterPrecentRange, MEGPrecentRange, TEGPrecentRange, ethanolPrecentRange, methanolPrecentRange;
        Excel.Range polarComPrecentRange;
        Excel.Range NaplusPrecentRange, ClminusPrecentRange;
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
            statusRange = this.Range["I22"];
            sumMolPrecentRange = this.Range["B118"];
            molPrecentCompRange = this.Range["B2", "B117"];
            allComponentRowsRange= this.Range["A1:A117"].EntireRow;
            extendedComponentRowsRange = this.Range["A44:A103"].EntireRow;
            extendedComponentRowsRange2 = this.Range["A109:A117"].EntireRow;
            extendedComponentRowsRange3 = this.Range["A10"].EntireRow;
            extendedComponentRowsRange4 = this.Range["A13:A17"].EntireRow;
            extendedComponentRowsRange5 = this.Range["A23"].EntireRow;
            extendedComponentRowsRange6 = this.Range["A26"].EntireRow;
            averageMolWtRange = this.Range["C118"];
            TBPCompRange = this.Range["B29", "B103"];
            waterPrecentRange = this.Range["B104"];
            MEGPrecentRange = this.Range["B107"];
            TEGPrecentRange = this.Range["B108"];
            ethanolPrecentRange = this.Range["B106"];
            methanolPrecentRange = this.Range["B105"];
            polarComPrecentRange = this.Range["B104", "B108"]; //???? check!
            NaplusPrecentRange = this.Range["B116"];
            ClminusPrecentRange = this.Range["B117"];
            EoScombobox.SelectedIndex = 0;
            inhibitorCalcTypecomboBox.SelectedIndex = 0;
            inhibitorComboBox.SelectedIndex = 0;
            numberOfPseudoCompComboBox.SelectedIndex = 0;
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.normalizeButton.Click += new System.EventHandler(this.normalizeButton_Click);
            this.clearButton.Click += new System.EventHandler(this.clearButton_Click);
            this.addComponentButoon.Click += new System.EventHandler(this.addComponentButoon_Click);
            this.molToWtButton.Click += new System.EventHandler(this.molToWtButton_Click);
            this.molRadioButton.Click += new System.EventHandler(this.molClicked);
            this.wtRadioButton.CheckedChanged += new System.EventHandler(this.wtRadioButton_CheckedChanged_1);
            this.wtRadioButton.Click += new System.EventHandler(this.wtClicked);
            this.plusFracRadioButton.Click += new System.EventHandler(this.plusClicked);
            this.noPlusRadioButton.Click += new System.EventHandler(this.noPlusClicked);
            this.tuneOkbutton.Click += new System.EventHandler(this.tuneOkbutton_Click);
            this.EoScombobox.SelectedIndexChanged += new System.EventHandler(this.EoScombobox_SelectedIndexChanged);
            this.button2.Click += new System.EventHandler(this.button2_Click);
            this.numberOfPseudoCompComboBox.SelectedIndexChanged += new System.EventHandler(this.numberOfPseudoCompComboBox_SelectedIndexChanged);
            this.inhibitorCalcTypecomboBox.SelectedIndexChanged += new System.EventHandler(this.inhibitorCalcTypecomboBox_SelectedIndexChanged);
            this.chemicalReactionsCheckBox.CheckedChanged += new System.EventHandler(this.chemicalReactionsCheckBox_CheckedChanged);
            this.dataBaseCheckBox.CheckedChanged += new System.EventHandler(this.dataBase_CheckedChanged);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
          
            if (sumMolPrecentRange.Value2 < 1e-100)
            {
                statusRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                statusRange.Value2 = "no components added..please set composition";
                return;
            }


            bool changed = false;
            if (this.Range["B1"].Value2 == "wt%")
            {
                changed = true;
                molToWtButton_Click(sender, e);
            }
            statusRange.Value2 = "normalizing fluid....please wait...";
            normalizeButton_Click(sender, e);

            statusRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            statusRange.Value2 = "creating fluid....please wait...";
            Cursor previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            SystemInterface thermoSystem = (SystemInterface)new SystemSrkEos(298, 10);

            if (this.Range["B109"].Value2 > 0 && !(this.Range["B109"].Value2 == null)) thermoSystem.addComponent("oxygen", this.Range["B109"].Value2);
            if (this.Range["B110"].Value2 > 0 && !(this.Range["B110"].Value2 == null)) thermoSystem.addComponent("hydrogen", this.Range["B110"].Value2);
            if (this.Range["B111"].Value2 > 0 && !(this.Range["B111"].Value2 == null)) thermoSystem.addComponent("argon", this.Range["B111"].Value2);
            if (this.Range["B112"].Value2 > 0 && !(this.Range["B112"].Value2 == null)) thermoSystem.addComponent("helium", this.Range["B112"].Value2);
            if (this.Range["B113"].Value2 > 0 && !(this.Range["B113"].Value2 == null)) thermoSystem.addComponent("mercury", this.Range["B113"].Value2);
            if (this.Range["B114"].Value2 > 0 && !(this.Range["B114"].Value2 == null)) thermoSystem.addComponent("S8", this.Range["B114"].Value2);
            if (this.Range["B115"].Value2 > 0 && !(this.Range["B115"].Value2 == null)) thermoSystem.addComponent("MDEA", this.Range["B115"].Value2);

            if (this.Range["B2"].Value2 > 0 && !(this.Range["B2"].Value2 == null)) thermoSystem.addComponent("nitrogen", this.Range["B2"].Value2);
            if (this.Range["B3"].Value2 > 0 && !(this.Range["B3"].Value2 == null)) thermoSystem.addComponent("CO2", this.Range["B3"].Value2);
            if (this.Range["B4"].Value2 > 0 && !(this.Range["B4"].Value2 == null)) thermoSystem.addComponent("H2S", this.Range["B4"].Value2);

            if (this.Range["B5"].Value2 > 0 && !(this.Range["B5"].Value2 == null)) thermoSystem.addComponent("methane", this.Range["B5"].Value2);
            if (this.Range["B6"].Value2 > 0 && !(this.Range["B6"].Value2 == null)) thermoSystem.addComponent("ethane", this.Range["B6"].Value2);
            if (this.Range["B7"].Value2 > 0 && !(this.Range["B7"].Value2 == null)) thermoSystem.addComponent("propane", this.Range["B7"].Value2);
            if (this.Range["B8"].Value2 > 0 && !(this.Range["B8"].Value2 == null)) thermoSystem.addComponent("i-butane", this.Range["B8"].Value2);
            if (this.Range["B9"].Value2 > 0 && !(this.Range["B9"].Value2 == null)) thermoSystem.addComponent("n-butane", this.Range["B9"].Value2);
            if (this.Range["B10"].Value2 > 0 && !(this.Range["B10"].Value2 == null)) thermoSystem.addComponent("22-dim-C3", this.Range["B10"].Value2);
            if (this.Range["B11"].Value2 > 0 && !(this.Range["B11"].Value2 == null)) thermoSystem.addComponent("i-pentane", this.Range["B11"].Value2);
            if (this.Range["B12"].Value2 > 0 && !(this.Range["B12"].Value2 == null)) thermoSystem.addComponent("n-pentane", this.Range["B12"].Value2);
            if (this.Range["B13"].Value2 > 0 && !(this.Range["B13"].Value2 == null)) thermoSystem.addComponent("c-C5", this.Range["B13"].Value2);
            if (this.Range["B14"].Value2 > 0 && !(this.Range["B14"].Value2 == null)) thermoSystem.addComponent("22-dim-C4", this.Range["B14"].Value2);
            if (this.Range["B15"].Value2 > 0 && !(this.Range["B15"].Value2 == null)) thermoSystem.addComponent("23-dim-C4", this.Range["B15"].Value2);
            if (this.Range["B16"].Value2 > 0 && !(this.Range["B16"].Value2 == null)) thermoSystem.addComponent("2-m-C5", this.Range["B16"].Value2);
            if (this.Range["B17"].Value2 > 0 && !(this.Range["B17"].Value2 == null)) thermoSystem.addComponent("3-m-C5", this.Range["B17"].Value2);


            if (this.Range["B18"].Value2 > 0 && !(this.Range["B18"].Value2 == null)) thermoSystem.addComponent("n-hexane", this.Range["B18"].Value2);
            if (this.Range["B19"].Value2 > 0 && !(this.Range["B19"].Value2 == null)) thermoSystem.addComponent("c-hexane", this.Range["B19"].Value2);
            if (this.Range["B20"].Value2 > 0 && !(this.Range["B20"].Value2 == null)) thermoSystem.addComponent("benzene", this.Range["B20"].Value2);
            if (this.Range["B21"].Value2 > 0 && !(this.Range["B21"].Value2 == null)) thermoSystem.addComponent("n-heptane", this.Range["B21"].Value2);
            if (this.Range["B22"].Value2 > 0 && !(this.Range["B22"].Value2 == null)) thermoSystem.addComponent("toluene", this.Range["B22"].Value2);
            if (this.Range["B23"].Value2 > 0 && !(this.Range["B23"].Value2 == null)) thermoSystem.addComponent("c-C7", this.Range["B23"].Value2);
            if (this.Range["B24"].Value2 > 0 && !(this.Range["B24"].Value2 == null)) thermoSystem.addComponent("n-octane", this.Range["B24"].Value2);
            if (this.Range["B25"].Value2 > 0 && !(this.Range["B25"].Value2 == null)) thermoSystem.addComponent("m-Xylene", this.Range["B25"].Value2);
            if (this.Range["B26"].Value2 > 0 && !(this.Range["B26"].Value2 == null)) thermoSystem.addComponent("c-C8", this.Range["B26"].Value2);
            if (this.Range["B27"].Value2 > 0 && !(this.Range["B27"].Value2 == null)) thermoSystem.addComponent("n-nonane", this.Range["B27"].Value2);
            if (this.Range["B28"].Value2 > 0 && !(this.Range["B28"].Value2 == null)) thermoSystem.addComponent("nC10", this.Range["B28"].Value2);

            int Cnumb = 5;
            for (int i = 29; i <= 103; i++)
            {
                Cnumb++;
                if (this.Range["B" + i.ToString()].Value2 > 0 && !(this.Range["B" + i.ToString()].Value2 == null)) thermoSystem.addTBPfraction(this.Range["A" + i.ToString()].Value2, this.Range["B" + i.ToString()].Value2, this.Range["C" + i.ToString()].Value2 / 1000.0, this.Range["D" + i.ToString()].Value2);
//                if (this.Range["B" + i.ToString()].Value2 > 0 && !(this.Range["B" + i.ToString()].Value2 == null)) thermoSystem.addTBPfraction("C" + Cnumb.ToString(), this.Range["B" + i.ToString()].Value2, this.Range["C" + i.ToString()].Value2 / 1000.0, this.Range["D" + i.ToString()].Value2);

            }

            if (plusFracRadioButton.Checked == true)
            {
                bool testIfTBP = thermoSystem.setHeavyTBPfractionAsPlusFraction();
                if (testIfTBP)
                {
                    thermoSystem.getCharacterization().getLumpingModel().setNumberOfPseudoComponents(Int16.Parse(numberOfPseudoCompComboBox.SelectedItem.ToString()));
                    thermoSystem.getCharacterization().characterisePlusFraction();
                }
            }

            if (waterPrecentRange.Value2 > 0 && !(waterPrecentRange.Value2 == null)) thermoSystem.addComponent("water", waterPrecentRange.Value2);
            if (methanolPrecentRange.Value2 > 0 && !(methanolPrecentRange.Value2 == null)) thermoSystem.addComponent("methanol", methanolPrecentRange.Value2);
            if (ethanolPrecentRange.Value2 > 0 && !(ethanolPrecentRange.Value2 == null)) thermoSystem.addComponent("ethanol", ethanolPrecentRange.Value2);
            if (MEGPrecentRange.Value2 > 0 && !(MEGPrecentRange.Value2 == null)) thermoSystem.addComponent("MEG", MEGPrecentRange.Value2);
            if (TEGPrecentRange.Value2 > 0 && !(TEGPrecentRange.Value2 == null)) thermoSystem.addComponent("TEG", TEGPrecentRange.Value2);
            if (NaplusPrecentRange.Value2 > 0 && !(NaplusPrecentRange.Value2 == null)) thermoSystem.addComponent("Na+", NaplusPrecentRange.Value2);
            if (ClminusPrecentRange.Value2 > 0 && !(ClminusPrecentRange.Value2 == null)) thermoSystem.addComponent("Cl-", ClminusPrecentRange.Value2);

            if (chemicalReactionsCheckBox.Checked)
            {
                thermoSystem.chemicalReactionInit();
            }

            thermoSystem.createDatabase(true);
            thermoSystem.setMixingRule(2);
            thermoSystem.init(0);
            //thermoSystem = thermoSystem.autoSelectModel();
            if (forceUseCPAmodel)
            {
                thermoSystem = thermoSystem.setModel("CPAs-SRK-EOS-statoil");
            }
            else if (EoScombobox.SelectedItem.ToString().Equals("Automatic"))
            {
                thermoSystem = thermoSystem.autoSelectModel();
            }
            else
            {
                thermoSystem = thermoSystem.setModel(EoScombobox.SelectedItem.ToString());
            }
            //thermoSystem.autoSelectMixingRule();

            if (iceCheckBox.Checked == true)
            {
                thermoSystem.setSolidPhaseCheck("water");
            }
            if (solidCO2checkBox.Checked == true)
            {
                thermoSystem.setSolidPhaseCheck("CO2");
            }
            if (solidGlycolCheckBox.Checked == true)
            {
                thermoSystem.setSolidPhaseCheck("MEG");
                thermoSystem.setSolidPhaseCheck("TEG");
            }
            if (waxcheckBox.Checked == true)
            {
                thermoSystem.getWaxModel().addTBPWax();
                thermoSystem.createDatabase(true);
                thermoSystem.addSolidComplexPhase("wax");
                thermoSystem.setMultiphaseWaxCheck(true);
            }
            if (chemicalReactionsCheckBox.Checked)
            {
                thermoSystem.chemicalReactionInit();
            }
            thermoSystem.autoSelectMixingRule();
            thermoSystem.setMultiPhaseCheck(true);
            thermoSystem.useVolumeCorrection(true);
            thermoSystem.init(0);
            thermoSystem.init(1);
            thermoSystem.initPhysicalProperties();
            //  thermoSystem.autoSelectMixingRule();
            //  Excel.Range range = this.Range["A7"];
            //   range.Value2 = "4.0";
            //   range.Calculate();

            statusRange.Value2 = "finished creating fluid";
            Cursor.Current = previousCursor;
            statusRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
            //    ThermodynamicOperations ops = new ThermodynamicOperations(thermoSystem);
            //  ops.TPflash();
            NeqSimThermoSystem.setThermoSystem(thermoSystem);

            if (changed)
            {
                molToWtButton_Click(sender, e);
            }

            Globals.Sheet9.button1_Click_1(sender, e);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            allComponentRowsRange.Select();
            allComponentRowsRange.Hidden = false;
        }

        private void addComponentButoon_Click(object sender, EventArgs e)
        {
            if (extendedComponentRowsRange.Hidden == true)
            {
                addComponentButoon.Text = "Hide Comps";
                extendedComponentRowsRange.Hidden = false;
                extendedComponentRowsRange2.Hidden = false;
                extendedComponentRowsRange3.Hidden = false;
                extendedComponentRowsRange4.Hidden = false;
                extendedComponentRowsRange5.Hidden = false;
                extendedComponentRowsRange6.Hidden = false;
            }
            else
            {
                addComponentButoon.Text = "Show More Comps";
                extendedComponentRowsRange.Hidden = true;
                extendedComponentRowsRange2.Hidden = true;
                extendedComponentRowsRange3.Hidden = true;
                extendedComponentRowsRange4.Hidden = true;
                extendedComponentRowsRange5.Hidden = true;
                extendedComponentRowsRange6.Hidden = true;
            }
        }

        private void normalizeButton_Click(object sender, EventArgs e)
        {
    
            int number = 1;
            double sumMolPrecent = sumMolPrecentRange.Value2;
            foreach (Excel.Range r in molPrecentCompRange.Cells)
            {
                number++;
                string textVar = "B" + number.ToString();
                if(!(this.Range[textVar].Value2 == null)){
                double newVal = this.Range[textVar].Value2 / sumMolPrecent * 100;
                this.Range[textVar].Value2 = newVal;
                }
            }
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            int number = 1;
            foreach (Excel.Range r in molPrecentCompRange.Cells)
            {
                number++;
                string textVar = "B" + number.ToString();
                this.Range[textVar].Value2 = 0;
            }

            number = 21;
            int Cnumber = 6;
            foreach (Excel.Range r in TBPCompRange.Cells)
            {
                string textVar = "A" + number.ToString();
                this.Range[textVar].Value2 = "C" + Cnumber.ToString();
                Cnumber++;
                number++;
            }

        }

        private void molToWtButton_Click(object sender, EventArgs e)
        {
            var val = this.Range["B1"].Value2;


            if (this.Range["B1"].Value2 == "mol%")
            {
                int number = 1;
                double molwt = averageMolWtRange.Value2; ;
                foreach (Excel.Range r in molPrecentCompRange.Cells)
                {
                    number++;
                    string textVar = "B" + number.ToString();
                    string textVar2 = "C" + number.ToString();
                    double newVal = this.Range[textVar].Value2 * this.Range[textVar2].Value2 / molwt;
                    this.Range[textVar].Value2 = newVal;
                }
            }
            else
            {
                //Excel.Range range = this.Range["B2", "B39"];
                int number = 1;
                double totalMoles = 0;
                double[] moles = new double[molPrecentCompRange.Count];
                foreach (Excel.Range r in molPrecentCompRange.Cells)
                {
                    number++;
                    string textVar = "B" + number.ToString();
                    string textVar2 = "C" + number.ToString();
                    if (this.Range[textVar].Value2>0 && !(this.Range[textVar].Value2 == null))
                    {
                        moles[number - 2] = this.Range[textVar].Value2 / this.Range[textVar2].Value2;
                        totalMoles += moles[number - 2];
                    }
                }
                number = 1;
                foreach (Excel.Range r in molPrecentCompRange.Cells)
                {
                    number++;
                    string textVar = "B" + number.ToString();
                    this.Range[textVar].Value2 = moles[number - 2] / totalMoles*100;
                }
            }


            if (val.Equals("mol%"))
            {
                this.Range["B1"].Value2 = "wt%";
                molToWtButton.Text = "Weight to Mol";
                molRadioButton.Checked = false;
                wtRadioButton.Checked = true;
            }
            else
            {
                this.Range["B1"].Value2 = "mol%";
                molToWtButton.Text = "Mol to Weight";
                molRadioButton.Checked = true;
                wtRadioButton.Checked = false;
            }

        }

        private void molRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            molRadioButton.Checked = true;
            wtRadioButton.Checked = false;
        }

        private void wtRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            molRadioButton.Checked = false;
            wtRadioButton.Checked = true;
        }

        private void plusFracRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            noPlusRadioButton.Checked = false;
        }

        private void noPlusRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            plusFracRadioButton.Checked = false;
        }

        private void plusClicked(object sender, EventArgs e)
        {
            noPlusRadioButton.Checked = false;
        }

        private void EoScombobox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataBase_CheckedChanged(object sender, EventArgs e)
        {
            if (dataBaseCheckBox.Checked)
            {
                neqsim.util.database.NeqSimDataBase.setDataBaseType("mySQL");
                neqsim.util.database.NeqSimDataBase.setConnectionString("jdbc:mysql://neqsim.equinor.com:3307/neqsimthermodatabase?useSSL=false");
                neqsim.util.database.NeqSimDataBase.setCreateTemporaryTables(true);

             }
            else
            {
                neqsim.util.database.NeqSimDataBase.setDataBaseType("Derby");
                neqsim.util.database.NeqSimDataBase.setConnectionString("jdbc:derby:classpath:data/neqsimthermodatabase");
                neqsim.util.database.NeqSimDataBase.setCreateTemporaryTables(false);
            }
        }

        private void noPlusClicked(object sender, EventArgs e)
        {
            plusFracRadioButton.Checked = false;
        }

        private void molClicked(object sender, EventArgs e)
        {
            wtRadioButton.Checked = false;
            this.Range["B1"].Value2 = "mol%";
            molToWtButton.Text = "Mol to Weight";
        }

        private void wtClicked(object sender, EventArgs e)
        {
            molRadioButton.Checked = false;
            this.Range["B1"].Value2 = "wt%";
            molToWtButton.Text = "Weight to Mol";
        }

        private void chemicalReactionsCheckBox_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void wtRadioButton_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void tuneOkbutton_Click(object sender, EventArgs e)
        {
           
            if (MWcheckBox.Checked == true)
            {

                
                double sumPlus = 0.0;
                foreach (Excel.Range r in TBPCompRange.Cells)
                {
                    if (r.Value2 > 1e-100)
                    {
                        sumPlus += r.Value2;
                    }
                }
                if (sumPlus < 1e-100)
                {
                    statusRange.Value2 = "no plus component set - please add plus component and set Mw and density....";
                    MWcheckBox.Checked = false;
                    tuneOkbutton_Click(sender, e);
                    return;
                }
               

                if (watsatcheckBox.Checked == true && (string)EoScombobox.SelectedItem=="Automatic")
                {
                    forceUseCPAmodel = true;
                }
                int numb = 21;
                foreach (Excel.Range r in TBPCompRange.Cells)
                {
                    if (r.Value2 > 1e-100) plusMWcell = "C" + numb;
                    numb++;
                }


                foreach (Excel.Range r in polarComPrecentRange.Cells)
                {
                    if (r.Value2 > 1e-100)
                    {
                        r.Value2 = 0;
                        statusRange.Value2 = "no polar components allowed during plus characterization - setting it to zero....";
                    }
                }

                noPlusRadioButton.Checked = true;
                plusFracRadioButton.Checked = false;
                button1_Click(sender, e);
                statusRange.Value2 = "estimating Mw of plus fraction to fit at saturation....";
                NeqSimThermoSystem.getThermoSystem().setTemperature(this.Range["F11"].Value2 + 273.15);
                NeqSimThermoSystem.getThermoSystem().setPressure(this.Range["F12"].Value2);
                NeqSimThermoSystem.getThermoSystem().setHeavyTBPfractionAsPlusFraction();
                NeqSimThermoSystem.getThermoSystem().getCharacterization().getLumpingModel().setNumberOfLumpedComponents(Int16.Parse(numberOfPseudoCompComboBox.SelectedItem.ToString()));         
                NeqSimThermoSystem.getThermoSystem().createDatabase(true);
                NeqSimThermoSystem.getThermoSystem().setMixingRule(NeqSimThermoSystem.getThermoSystem().getMixingRule());
                NeqSimThermoSystem.getThermoSystem().init(0);
                NeqSimThermoSystem.getThermoSystem().init(1);

                SaturationPressureFunction function = new SaturationPressureFunction();
                double[] guess = { this.Range[plusMWcell].Value2/1000.0};
                function.setInitialGuess(guess);

                List<SampleValue> sampleList = new List<SampleValue>();
                double[] sample1 = { this.Range["F11"].Value2 + 273.15 };
                double satPres = this.Range["F12"].Value2;
                double[] standardDeviation1 = {1.5};
                SampleValue[] sample = new SampleValue[1];
                sample[0] = new SampleValue(satPres, satPres / 100.0, sample1, standardDeviation1);
                sample[0].setFunction(function);
                function.setInitialGuess(guess);
                sample[0].setThermodynamicSystem(NeqSimThermoSystem.getThermoSystem());
                SampleSet sampleSet = new SampleSet(sample);
                LevenbergMarquardt optim = new LevenbergMarquardt();
                optim.setMaxNumberOfIterations(20);
                optim.setSampleSet(sampleSet);
                optim.solve();
                statusRange.Value2 = "finished creating fluid";

                //double a = optim.getSample(0).getFunction().getFittingParams(0) * 1000.0;
                //range.Value2 = a;

                

                this.Range[plusMWcell].Value2 = optim.getSample(0).getFunction().getFittingParams(0) * 1000.0;
                noPlusRadioButton.Checked = false;
                plusFracRadioButton.Checked = true;
                
            }
            button1_Click(sender, e);
            if (watsatcheckBox.Checked == true)
            {
               // this.Range["B36"].Value2 = 5.0; //adding some water to the fluid

                 statusRange.Value2 = "estimating water content in gas at saturation....";
                 NeqSimThermoSystem.getThermoSystem().setTemperature(this.Range["F11"].Value2 + 273.15);
                 NeqSimThermoSystem.getThermoSystem().setPressure(this.Range["F12"].Value2);

                 NeqSimThermoSystem.getThermoSystem().addComponent("water", NeqSimThermoSystem.getThermoSystem().getTotalNumberOfMoles());
                 NeqSimThermoSystem.getThermoSystem().createDatabase(true);
                 NeqSimThermoSystem.getThermoSystem().setMixingRule(NeqSimThermoSystem.getThermoSystem().getMixingRule());
                 NeqSimThermoSystem.setThermoSystem(NeqSimThermoSystem.getThermoSystem().autoSelectModel());

                 ThermodynamicOperations ops = new ThermodynamicOperations(NeqSimThermoSystem.getThermoSystem());
                 ops.saturateWithWater();
                 int maxIter = 100, iter=0;
                 do
                 {
                     iter++;
                     waterPrecentRange.Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent("water").getz()*100;
                     normalizeButton_Click(sender, e);
                 }
                 while (iter < maxIter && Math.Abs((waterPrecentRange.Value2/100.0 - NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent("water").getx()) / NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent("water").getx()) > 1e-6);
                statusRange.Value2 = "finished creating fluid";

            
            }
            forceUseCPAmodel = false;
            Globals.Sheet9.button1_Click_1(sender, e);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (waterPrecentRange.Value2 == 0)
            {
                statusRange.Value2 = "please add water before estimating inhibitor concentration...";
                return;
            }
            bool changed = false;
            if (this.Range["B1"].Value2 == "wt%")
            {
                changed = true;
                molToWtButton_Click(sender, e);
            }

            Excel.Range inhibitorRange = MEGPrecentRange;
            if (inhibitorComboBox.SelectedItem.ToString().Equals("MEG")) inhibitorRange = MEGPrecentRange;
            else if (inhibitorComboBox.SelectedItem.ToString().Equals("TEG")) inhibitorRange = TEGPrecentRange;
            else if (inhibitorComboBox.SelectedItem.ToString().Equals("methanol")) inhibitorRange = methanolPrecentRange;
            else if (inhibitorComboBox.SelectedItem.ToString().Equals("ethanol")) inhibitorRange = ethanolPrecentRange;


            if (inhibitorCalcTypecomboBox.SelectedItem.ToString().Equals("estimate wt%"))
            {
               
                inhibitorRange.Value2 = waterPrecentRange.Value2 / 3.0;
                button1_Click(sender, e);
                statusRange.Value2 = "estimating inhibitor concentration...please wait";
                NeqSimThermoSystem.getThermoSystem().setHydrateCheck(true);
                ThermodynamicOperations ops = new ThermodynamicOperations(NeqSimThermoSystem.getThermoSystem());
                NeqSimThermoSystem.getThermoSystem().setPressure(this.Range["L12"].Value2);
                NeqSimThermoSystem.getThermoSystem().setTemperature(this.Range["L11"].Value2 + 273.15);
                string inhibitorname = inhibitorComboBox.SelectedItem.ToString();

                ops.hydrateInhibitorConcentration(inhibitorname, this.Range["L11"].Value2 + 273.15);

                double zInh = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(inhibitorname).getz();

                int maxIter = 100, iter = 0;
                double error = 0, oldz = 0.0;
                do
                {
                    iter++;
                    oldz = inhibitorRange.Value2;
                    inhibitorRange.Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(inhibitorname).getz() * 100;
                    normalizeButton_Click(sender, e);
                    error = Math.Abs(oldz - inhibitorRange.Value2);
                }
                while (error > 1e-4 && iter < maxIter);
                button1_Click(sender, e);
                Globals.Sheet9.button1_Click_1(sender, e);
            }

            else
            {
                inhibitorRange.Value2 = waterPrecentRange.Value2 / 3.0;
                button1_Click(sender, e);
                statusRange.Value2 = "adjusting inhibitor concentration to set point...please wait";

                ThermodynamicOperations ops = new ThermodynamicOperations(NeqSimThermoSystem.getThermoSystem());
                NeqSimThermoSystem.getThermoSystem().setPressure(this.Range["L12"].Value2);
                NeqSimThermoSystem.getThermoSystem().setTemperature(this.Range["L11"].Value2 + 273.15);
                string inhibitorname = inhibitorComboBox.SelectedItem.ToString();

                ops.hydrateInhibitorConcentrationSet(inhibitorname, this.Range["L13"].Value2/100.0);

                double zInh = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(inhibitorname).getz();

                int maxIter = 100, iter = 0;
                double error = 0, oldz = 0.0;
                do
                {
                    iter++;
                    oldz = inhibitorRange.Value2;
                    inhibitorRange.Value2 = NeqSimThermoSystem.getThermoSystem().getPhase(0).getComponent(inhibitorname).getz() * 100;
                    normalizeButton_Click(sender, e);
                    error = Math.Abs(oldz - inhibitorRange.Value2);
                }
                while (error > 1e-4 && iter < maxIter);
                button1_Click(sender, e);
                Globals.Sheet9.button1_Click_1(sender, e);


            }

            if (changed)
            {
                molToWtButton_Click(sender, e);
            }

        }

        private void numberOfPseudoCompComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Sheet1_Deactivate()
        {

        }

        private void inhibitorCalcTypecomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (inhibitorCalcTypecomboBox.SelectedItem.ToString().Equals("estimate wt%"))
            {
                this.Range["K11"].Value2 = "Thyd [C]";
                this.Range["K12"].Value2 = "Phyd [bara]";
                this.Range["K13"].Value2 = "";
                this.Range["L13"].Value2 = "";

            }
            else{
                this.Range["K11"].Value2 = "Temp [C]";
                this.Range["K12"].Value2 = "Pres [bara]";
                this.Range["K13"].Value2 = "wt%";
                this.Range["L13"].Value2 = "50";
                }
        }

        

    }
}
