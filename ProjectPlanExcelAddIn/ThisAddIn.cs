using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ProjectPlanExcelAddIn
{
    public partial class ThisAddIn
    {
        public GPTManager GPTManager { get; set; }
        public ProjectProductStorage _storage;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookActivate += new AppEvents_WorkbookActivateEventHandler(Application_WorkbookOpen);
            GPTManager = new GPTManager();

            _storage = new ProjectProductStorage();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.WorkbookOpen -= new AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            _storage.SaveToFile();
        }

        private void Application_WorkbookOpen(Workbook workbook)
        {
            // Выполняем проверку темы и настройку кнопок
            Globals.Ribbons.RibbonPlan.CheckTemplateAndConfigure();
        }

        public string GetTemplateID(Workbook workbook)
        {
            try
            {
                foreach (var property in workbook.CustomDocumentProperties)
                {
                    var prop = (Microsoft.Office.Core.DocumentProperty)property;
                    if (prop.Name == "TemplateID")
                    {
                        return prop.Value.ToString();
                    }
                }
            }
            catch
            {
                // Обработка ошибок
            }

            return "Unknown";
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