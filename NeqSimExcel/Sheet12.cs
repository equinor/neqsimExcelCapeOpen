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

namespace NeqSimExcel
{
    public partial class Sheet12
    {
        private void Sheet12_Startup(object sender, System.EventArgs e)
        {
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
            this.fieldTypeCombobox.SelectedIndexChanged += new System.EventHandler(this.fieldTypeCombobox_SelectedIndexChanged);
            this.Startup += new System.EventHandler(this.Sheet12_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet12_Shutdown);

        }

        #endregion

        private void fieldTypeCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
