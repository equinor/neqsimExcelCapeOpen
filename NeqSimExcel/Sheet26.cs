using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using thermo.system;
using thermodynamicOperations;

namespace NeqSimExcel
{
    public partial class Sheet26
    {
        private void Sheet26_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet26_Shutdown(object sender, System.EventArgs e)
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
            this.Startup += new System.EventHandler(this.Sheet26_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet26_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {

            SystemInterface thermoSystem = (SystemInterface)NeqSimThermoSystem.getThermoSystem();//.clone();
            thermoSystem.init(0);
            thermoSystem.init(1);

            int Cnumb = 5;
          for (int i = 2; i <= 100; i++)
          {
               string a =  this.Range["A" + i.ToString()].Value2;
                if(a!=null) thermoSystem.addTBPfraction(this.Range["A" + i.ToString()].Value2, 1.0e-10, (this.Range["C" + i.ToString()].Value2 / 1000.0), (this.Range["D" + i.ToString()].Value2 / 1000.0));
            }

            thermoSystem.autoSelectMixingRule();
            thermoSystem.init(0);
            thermoSystem.init(1);
            thermoSystem.setMultiPhaseCheck(true);
            thermoSystem.resetPhysicalProperties();
            thermoSystem.initPhysicalProperties();
            NeqSimThermoSystem.setThermoSystem(thermoSystem);
            Globals.Sheet9.button1_Click_1(sender, e);
        }
    }
}
