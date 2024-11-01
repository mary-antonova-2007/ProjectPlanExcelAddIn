﻿using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace ProjectPlanExcelAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookOpen += new AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
        }

        private void Application_WorkbookOpen(Workbook workbook)
        {
            // Выполняем проверку темы и настройку кнопок
            Globals.Ribbons.RibbonPlan.CheckTemplateAndConfigure();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.WorkbookOpen -= new AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
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