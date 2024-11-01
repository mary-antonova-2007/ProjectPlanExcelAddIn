using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ProjectPlanExcelAddIn
{
    public partial class RibbonPlan
    {
        private string TemplateID
        {
            get
            {
                return Globals.ThisAddIn.GetTemplateID(Globals.ThisAddIn.Application.ActiveWorkbook);
            }
        }
        private void RibbonPlan_Load(object sender, RibbonUIEventArgs e)
        {
        }

        public static void RunDateShifter()
        {
            // Создаем экземпляр пользовательской формы InputForm
            var inputForm = new InputForm { LabelInfo = "Введите количество дней для сдвига:" };

            // Открываем форму и ждем результата
            if (inputForm.ShowDialog() == DialogResult.OK)
            {
                if (int.TryParse(inputForm.TextBoxData, out int shiftDays))
                {
                    Application excelApp = Globals.ThisAddIn.Application;
                    Range selectedRange = excelApp.Selection as Range;

                    if (selectedRange != null)
                    {
                        var dateShifter = new DateShifter(new BusinessCalendar(), shiftDays);
                        dateShifter.ShiftSelectedDates(selectedRange);

                        MessageBox.Show("Даты успешно сдвинуты!", "Готово");
                    }
                    else
                    {
                        MessageBox.Show("Пожалуйста, выделите диапазон ячеек.", "Ошибка");
                    }
                }
                else
                {
                    MessageBox.Show("Введите корректное число для сдвига.", "Ошибка");
                }
            }
            else
            {
                // Обработка отмены операции
                MessageBox.Show("Операция отменена пользователем.", "Отмена");
            }
        }

        public void CheckTemplateAndConfigure()
        {
            Application excelApp = Globals.ThisAddIn.Application;
            Workbook activeWorkbook = excelApp.ActiveWorkbook;

            if (activeWorkbook == null)
            {
                MessageBox.Show("Нет открытой книги для проверки шаблона.", "Ошибка");
                return;
            }

            // Получаем тему книги
            string workbookTheme = GetWorkbookTheme(activeWorkbook);

            // Настраиваем функционал в зависимости от темы
            if (workbookTheme == "plan")
            {
                EnablePlanFunctions();
                MessageBox.Show("Планер активен");
            }
            else
            {
                DisableAllFunctions();
                MessageBox.Show("Тема книги не распознана. Функционал ограничен.", "Предупреждение");
            }
        }

        private string GetWorkbookTheme(Workbook workbook)
        {
            try
            {
                // Получаем значение свойства "Тема" из документа
                return workbook.BuiltinDocumentProperties["Subject"].Value.ToString();
            }
            catch
            {
                return string.Empty; // Если тема не задана, возвращаем пустую строку
            }
        }

        private void EnablePlanFunctions()
        {
            // Включаем кнопки и функции, связанные с темой "plan"

        }

        private void DisableAllFunctions()
        {
            // Отключаем все кнопки и функции

        }

        private void buttonCheckTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            CheckTemplateAndConfigure();
        }
        private void buttonAddDays_Click(object sender, RibbonControlEventArgs e)
        {
            RunDateShifter();
        }

        private void buttonAutoPlan_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Получаем экземпляр Excel Application и активный лист
                Application excelApp = Globals.ThisAddIn.Application;
                Worksheet activeWorksheet = excelApp.ActiveSheet as Worksheet;

                if (activeWorksheet != null)
                {
                    // Создаем экземпляр календаря, который будет учитывать рабочие и праздничные дни
                    BusinessCalendar businessCalendar = new BusinessCalendar();

                    // Создаем экземпляр планировщика задач и запускаем планирование
                    TaskScheduler taskScheduler = new TaskScheduler(excelApp);
                    taskScheduler.ScheduleTasks(activeWorksheet);

                    MessageBox.Show("Планирование задач завершено успешно!", "Готово");
                }
                else
                {
                    MessageBox.Show("Не удалось получить активный лист. Пожалуйста, откройте таблицу с задачами.", "Ошибка");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при планировании задач: {ex.Message}", "Ошибка");
            }
        }
    }
}
