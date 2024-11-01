using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;

namespace ProjectPlanExcelAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Сохраняем календарь перед завершением работы

        }

        #region Код, созданный VSTO

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}