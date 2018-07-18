using System;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class RecombineSheet
    {
        private void Sheet19_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet19_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += Sheet19_Startup;
            Shutdown += Sheet19_Shutdown;
        }

        #endregion
    }
}