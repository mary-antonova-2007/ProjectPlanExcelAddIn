using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Threading.Tasks;
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
        Application ExcelApp {
            get {
                return Globals.ThisAddIn.Application;
            }
        }
        private void RibbonPlan_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void RowsDateShifter(int countDays)
        {
            // Получаем ссылку на приложение Excel
            Application excelApp = Globals.ThisAddIn.Application;
            Range selectedRange = excelApp.Selection as Range;

            // Проверяем, что диапазон выделен
            if (selectedRange != null)
            {
                // Создаем экземпляр DateShifter с переданным значением для сдвига
                var dateShifter = new DateShifter(new BusinessCalendar(), countDays);

                // Получаем номер первой строки выделенного диапазона
                int firstRow = selectedRange.Row;

                // Получаем номер последней строки выделенного диапазона
                int lastRow = selectedRange.Row + selectedRange.Rows.Count - 1;

                // Проходим по каждой строке в выделенном диапазоне
                for (int row = firstRow; row <= lastRow; row++)
                {
                    // Получаем всю строку, в которой находится выделенная ячейка
                    Range entireRow = selectedRange.Worksheet.Rows[row];

                    // Сдвигаем даты в текущей строке
                    dateShifter.ShiftDatesInRows(entireRow);
                }

                // Выводим всплывающее уведомление в панель состояния Excel
                ShowMessageInStatusBar("Даты в строках выделенного диапазона успешно сдвинуты.");
                ShowNotification("Даты в строках выделенного диапазона успешно сдвинуты.");
            }
            else
            {
                // Если диапазон не выбран, выводим сообщение в панель состояния
                ShowMessageInStatusBar("Пожалуйста, выделите диапазон ячеек.");
                ShowNotification("Пожалуйста, выделите диапазон ячеек.");
            }
        }

        private void ShiftDatesLeft(object sender, RibbonControlEventArgs e)
        {
            RowsDateShifter(-1);
        }

        private void ShiftDatesRight(object sender, RibbonControlEventArgs e)
        {
            RowsDateShifter(1);
        }

        private void RunDateShifter(object sender, RibbonControlEventArgs e)
        {
            // Создаем экземпляр пользовательской формы InputForm
            var inputForm = new InputForm { LabelInfo = "Введите количество дней для сдвига:" };

            // Открываем форму и ждем результата
            if (inputForm.ShowDialog() == DialogResult.OK)
            {
                if (int.TryParse(inputForm.TextBoxData, out int shiftDays))
                {
                    Range selectedRange = ExcelApp.Selection as Range;

                    if (selectedRange != null)
                    {
                        var dateShifter = new DateShifter(new BusinessCalendar(), shiftDays);
                        dateShifter.ShiftSelectedDates(selectedRange);
                        ShowMessageInStatusBar("Даты успешно сдвинуты!");
                    }
                    else
                    {
                        ShowMessageInStatusBar("Пожалуйста, выделите диапазон ячеек.");
                    }
                }
                else
                {
                    ShowMessageInStatusBar("Введите корректное число для сдвига.");
                }
            }
            else
            {
                // Обработка отмены операции
                ShowMessageInStatusBar("Операция отменена пользователем.");

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
            }
            else
            {
                DisableAllFunctions();
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
            groupPlan.Visible = true;
        }

        private void DisableAllFunctions()
        {
            // Отключаем все кнопки и функции
            groupPlan.Visible = false;
        }

        private void buttonCheckTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            CheckTemplateAndConfigure();
        }

        private void buttonAddDays_Click(object sender, RibbonControlEventArgs e)
        {
            RunDateShifter(sender, e);
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
        
        private async void buttonGPTQuestion_Click(object sender, RibbonControlEventArgs e)
        {
            var manager = Globals.ThisAddIn.GPTManager;

            // Получение текста из выбранной ячейки как запроса
            string prompt = Globals.ThisAddIn.Application.ActiveCell.Text;
            string selectedRange = Globals.ThisAddIn.Application.Selection.Address;

            // Отправка запроса и получение ответа
            string response = await manager.GetResponseAsync(prompt, selectedRange);

            // Выполнение команд, если ответ содержит JSON-команды
            if (!string.IsNullOrWhiteSpace(response))
            {
                manager.ExecuteCommands(response);
            }
        }

        private void buttonGPTSettings_Click(object sender, RibbonControlEventArgs e)
        {
            GPTSettingsForm form = new GPTSettingsForm();
            form.ShowDialog();
        }

        public void MoveRowsUp(object sender, RibbonControlEventArgs e)
        {
            var excel = Globals.ThisAddIn.Application;
            var sheet = excel.ActiveSheet as Worksheet;
            var selection = excel.Selection as Range;

            if (selection == null) return;

            Range rows = selection.EntireRow;
            int firstRowIndex = rows.Row;
            int rowCount = rows.Rows.Count;

            if (firstRowIndex <= 1) return; // Нельзя двигать выше первой строки

            Range aboveRows = sheet.Rows[firstRowIndex - 1];

            // Вырезаем выделенные строки
            rows.Cut();
            aboveRows.Insert(XlInsertShiftDirection.xlShiftDown);

            // Обновляем выделение, выделяя строки на новом месте
            sheet.Range[sheet.Rows[firstRowIndex - 1], sheet.Rows[firstRowIndex + rowCount - 2]].Select();
        }

        public void MoveRowsDown(object sender, RibbonControlEventArgs e)
        {
            var excel = Globals.ThisAddIn.Application;
            var sheet = excel.ActiveSheet as Worksheet;
            var selection = excel.Selection as Range;

            if (selection == null) return;

            Range rows = selection.EntireRow;
            int firstRowIndex = rows.Row;
            int rowCount = rows.Rows.Count;
            int lastRow = sheet.Rows.Count;

            if (firstRowIndex + rowCount > lastRow) return; // Если строки внизу нет, не двигаем

            Range belowRows = sheet.Rows[firstRowIndex + rowCount];

            // Вырезаем и вставляем строки
            belowRows.Cut();
            rows.Insert(XlInsertShiftDirection.xlShiftDown);

            // Обновляем выделение на новые строки
            sheet.Range[sheet.Rows[firstRowIndex + 1], sheet.Rows[firstRowIndex + rowCount]].Select();
        }

        public void InsertRowAbove(object sender, RibbonControlEventArgs e)
        {
            var excel = Globals.ThisAddIn.Application;
            var sheet = excel.ActiveSheet as Worksheet;
            var selection = excel.Selection as Range;

            if (selection == null) return;

            Range row = selection.EntireRow;
            int rowIndex = row.Row;

            // Отключаем обновление экрана и вычисления для ускорения
            excel.ScreenUpdating = false;
            excel.Calculation = XlCalculation.xlCalculationManual;

            // Копируем строку и вставляем выше
            row.Copy();

            // Вставляем строку выше текущей строки
            Range newRow = sheet.Rows[rowIndex]; // Указываем строку, которая будет вставлена выше
            newRow.Insert(XlInsertShiftDirection.xlShiftDown);  // Вставляем строку

            // Очищаем все значения в новой строке, но оставляем формулы
            try
            {
                Range constants = newRow.SpecialCells(XlCellType.xlCellTypeConstants);
                if (constants != null)
                {
                    constants.ClearContents();
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Если нет значений, просто пропускаем
            }

            // Обновляем выделение
            newRow.Select();

            // Включаем обновление экрана и вычисления обратно
            excel.ScreenUpdating = true;
            excel.Calculation = XlCalculation.xlCalculationAutomatic;
            MoveRowsUp(sender, e);
        }

        private void ShowNotification(string message, string title = "Уведомление")
        {
            NotifyIcon notifyIcon = new NotifyIcon();
            notifyIcon.Icon = SystemIcons.Information;
            notifyIcon.Visible = true;
            notifyIcon.BalloonTipText = message;
            notifyIcon.BalloonTipTitle = title;
            notifyIcon.ShowBalloonTip(1500); // Покажет уведомление на 1.5 секунды
        }

        private void ShowMessageInStatusBar(string message, int pauseMillisec = 3000)
        {
            ExcelApp.StatusBar = message;
            System.Threading.Thread.Sleep(pauseMillisec); // Задержка pauseMillisec милисекунды
            ExcelApp.StatusBar = false; // Сбрасываем панель состояния
        }

    }
}
