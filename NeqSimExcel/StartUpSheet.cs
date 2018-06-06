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
using System.Reflection;
namespace NeqSimExcel
{
    public partial class Sheet12
    {
        private void Sheet12_Startup(object sender, System.EventArgs e)
        {
            Globals.PVTSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet15.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Olga.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet6.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet17.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.hydrate.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet10.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet11.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet18.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.RecombineSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.MixSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet5.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet13.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet14.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet2.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet11.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.valveSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.compressorSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.SepProcessSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet25.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet24.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet26.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet27.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet28.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.Sheet29.Visible = Excel.XlSheetVisibility.xlSheetHidden;

        }

        private void Sheet12_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.activateOperationsCheckBox.CheckedChanged += new System.EventHandler(this.activateOperationsCheckBox_CheckedChanged);
            this.activatePVTcheckBox.CheckedChanged += new System.EventHandler(this.activatePVTcheckBox_CheckedChanged);
            this.exportCheckBox.CheckedChanged += new System.EventHandler(this.exportCheckBox_CheckedChanged);
            this.RFOCheckBox.CheckedChanged += new System.EventHandler(this.RFOCheckBox_CheckedChanged);
            this.activateFACheckBox.CheckedChanged += new System.EventHandler(this.activateFACheckBox_CheckedChanged);
            this.advancedCheckBox.CheckedChanged += new System.EventHandler(this.advancedCheckBox_CheckedChanged);
            this.activateFluidOperationCheckBox.CheckedChanged += new System.EventHandler(this.activateFluidOperationCheckBox_CheckedChanged);
            this.unitOperationcheckBox.CheckedChanged += new System.EventHandler(this.unitOperationcheckBox_CheckedChanged);
            this.processCheckBox.CheckedChanged += new System.EventHandler(this.processCheckBox_CheckedChanged);
            this.Startup += new System.EventHandler(this.Sheet12_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet12_Shutdown);

        }

        #endregion

        private void activateOperationsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (activateOperationsCheckBox.Checked)
            {
                Globals.Sheet3.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet5.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet13.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet14.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet2.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet11.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet29.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            }
            else
            {
                Globals.Sheet3.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet5.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet13.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet14.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet2.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet11.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet29.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            }
        }

        private void activatePVTcheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (activatePVTcheckBox.Checked)
            {
                Globals.PVTSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet15.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.PVTSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet15.Visible = Excel.XlSheetVisibility.xlSheetHidden;

            }
        }

        private void exportCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (exportCheckBox.Checked)
            {
                Globals.Olga.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet6.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet27.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.Olga.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet6.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet27.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            }
        }

        private void RFOCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RFOCheckBox.Checked)
            {
                Globals.Sheet17.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.Sheet17.Visible = Excel.XlSheetVisibility.xlSheetHidden;

            }
        }

        private void activateFACheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (activateFACheckBox.Checked)
            {
                Globals.hydrate.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet10.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet14.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.hydrate.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet10.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet14.Visible = Excel.XlSheetVisibility.xlSheetHidden;

            }
        }

        private void advancedCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (advancedCheckBox.Checked)
            {
                Globals.Sheet18.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet24.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.Sheet18.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet24.Visible = Excel.XlSheetVisibility.xlSheetHidden;

            }
        }

        private void activateFluidOperationCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (activateFluidOperationCheckBox.Checked)
            {
                Globals.RecombineSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.MixSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet15.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet26.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet28.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.RecombineSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.MixSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet15.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet26.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet28.Visible = Excel.XlSheetVisibility.xlSheetHidden;

            }
        }

        private void unitOperationcheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (unitOperationcheckBox.Checked)
            {
                Globals.valveSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.compressorSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
           
            }
            else
            {
                Globals.valveSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.compressorSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
           
            }
        }

        private void processCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (processCheckBox.Checked)
            {
                Globals.SepProcessSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.Sheet25.Visible = Excel.XlSheetVisibility.xlSheetVisible;
           }
            else
            {
                Globals.SepProcessSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                Globals.Sheet25.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            }
        }

      

   
    

    }
}
