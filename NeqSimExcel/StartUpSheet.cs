using System;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class Sheet12
    {
        private void Sheet12_Startup(object sender, EventArgs e)
        {
            Globals.PVTSheet.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet15.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Olga.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet6.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet17.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.hydrate.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet10.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet11.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet18.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.RecombineSheet.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.MixSheet.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet5.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet13.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet14.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet2.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet11.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.valveSheet.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.compressorSheet.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.SepProcessSheet.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet25.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet24.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet26.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet27.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet28.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet29.Visible = XlSheetVisibility.xlSheetHidden;
            Globals.Sheet30.Visible = XlSheetVisibility.xlSheetHidden;
        }

        private void Sheet12_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
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
                Globals.Sheet3.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet5.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet13.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet29.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet2.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet11.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet29.Visible = XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.Sheet3.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet5.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet13.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet29.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet2.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet11.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet29.Visible = XlSheetVisibility.xlSheetHidden;
            }

        }

        private void activatePVTcheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (activatePVTcheckBox.Checked)
            {
                Globals.PVTSheet.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet15.Visible = XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.PVTSheet.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet15.Visible = XlSheetVisibility.xlSheetHidden;
            }
        }

        private void exportCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (exportCheckBox.Checked)
            {
                Globals.Olga.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet6.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet27.Visible = XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.Olga.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet6.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet27.Visible = XlSheetVisibility.xlSheetHidden;
            }
        }

        private void RFOCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RFOCheckBox.Checked)
                Globals.Sheet17.Visible = XlSheetVisibility.xlSheetVisible;
            else
                Globals.Sheet17.Visible = XlSheetVisibility.xlSheetHidden;
        }

        private void activateFACheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (activateFACheckBox.Checked)
            {
                Globals.hydrate.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet10.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet14.Visible = XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.hydrate.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet10.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet14.Visible = XlSheetVisibility.xlSheetHidden;
            }
        }

        private void advancedCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (advancedCheckBox.Checked)
            {
                Globals.Sheet18.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet24.Visible = XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.Sheet18.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet24.Visible = XlSheetVisibility.xlSheetHidden;
            }
        }

        private void activateFluidOperationCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (activateFluidOperationCheckBox.Checked)
            {
                Globals.RecombineSheet.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.MixSheet.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet15.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet26.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet28.Visible = XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.RecombineSheet.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.MixSheet.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet15.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet26.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet28.Visible = XlSheetVisibility.xlSheetHidden;
            }
        }

        private void unitOperationcheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (unitOperationcheckBox.Checked)
            {
                Globals.valveSheet.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.compressorSheet.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet30.Visible = XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.valveSheet.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.compressorSheet.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet30.Visible = XlSheetVisibility.xlSheetHidden;
            }
        }

        private void processCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (processCheckBox.Checked)
            {
                Globals.SepProcessSheet.Visible = XlSheetVisibility.xlSheetVisible;
                Globals.Sheet25.Visible = XlSheetVisibility.xlSheetVisible;
            }
            else
            {
                Globals.SepProcessSheet.Visible = XlSheetVisibility.xlSheetHidden;
                Globals.Sheet25.Visible = XlSheetVisibility.xlSheetHidden;
            }
        }
    }
}