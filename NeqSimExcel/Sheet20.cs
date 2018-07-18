using System;
using Office = Microsoft.Office.Core;

namespace NeqSimExcel
{
    public partial class MixSheet
    {
        private void Sheet20_Startup(object sender, EventArgs e)
        {
        }

        private void Sheet20_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += Sheet20_Startup;
            Shutdown += Sheet20_Shutdown;
        }

        #endregion
    }
}