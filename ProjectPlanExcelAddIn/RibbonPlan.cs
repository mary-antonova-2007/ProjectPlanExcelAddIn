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
using System.Globalization;
using System.Linq;
using System.Text;

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
        string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ProjectPlan", "Data");
        ProjectProductStorage ProjectProductStorage { get
            {
                return Globals.ThisAddIn._storage;
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
            }
            else
            {
                // Если диапазон не выбран, выводим сообщение в панель состояния
                ShowMessageInStatusBar("Пожалуйста, выделите диапазон ячеек.");
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

        private void RunDateShifterWithForm(object sender, RibbonControlEventArgs e)
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
            RunDateShifterWithForm(sender, e);
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
        }

        private void buttonProductTimeReport_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var files = PromptUserToSelectFiles();
            if (files == null || files.Length == 0) return;

            var taskData = new Dictionary<string, Dictionary<string, double>>(StringComparer.OrdinalIgnoreCase);
            var columnMeta = new SortedSet<(int year, int month, string user, string title)>();

            foreach (string file in files)
            {
                ProcessFile(file, app, taskData, columnMeta);
            }

            GenerateReportSheet(app, taskData, columnMeta);
            MessageBox.Show("Сводный отчет успешно создан!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string[] PromptUserToSelectFiles()
        {
            using (var dialog = new OpenFileDialog
            {
                Title = "Выберите Excel-файлы",
                Multiselect = true,
                Filter = "Excel файлы (*.xlsx)|*.xlsx"
            })
            {
                return dialog.ShowDialog() == DialogResult.OK ? dialog.FileNames : null;
            }
        }

        private string[] PromptUserToSelectFolderAndGetFiles()
        {
            using (var dialog = new FolderBrowserDialog
            {
                Description = "Выберите папку с Excel-файлами"
            })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string folder = dialog.SelectedPath;
                    var files = Directory.GetFiles(folder, "*.xlsx", SearchOption.AllDirectories);
                    return files;
                }
            }
            return null;
        }

        private void buttonCreateReport_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var files = PromptUserToSelectFolderAndGetFiles();
            if (files == null || files.Length == 0)
            {
                MessageBox.Show("Файлы не найдены.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var taskData = new Dictionary<string, Dictionary<string, double>>(StringComparer.OrdinalIgnoreCase);
            var columnMeta = new SortedSet<(int year, int month, string user, string title)>();

            foreach (string file in files)
            {
                if (!file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) continue;
                ProcessFile(file, app, taskData, columnMeta);
            }

            GenerateReportSheet(app, taskData, columnMeta);
            MessageBox.Show("Сводный отчет успешно создан!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string SanitizeSheetName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars().Concat(new[] { '[', ']', '*', '?', '/', '\\' });
            foreach (var c in invalid)
                name = name.Replace(c, '_');
            return name.Length > 31 ? name.Substring(0, 31) : name;
        }

        private void GenerateReportSheet(Application app, Dictionary<string, Dictionary<string, double>> taskData, SortedSet<(int year, int month, string user, string title)> columnMeta)
        {
            var groupedByUser = columnMeta.GroupBy(x => x.user);

            foreach (var userGroup in groupedByUser)
            {
                string user = userGroup.Key;

                var ws = app.ActiveWorkbook.Sheets.Add();
                ws.Name = SanitizeSheetName(user);

                // Сформируем мапу: ключ = $"{year}_{month}", значение = индекс столбца
                var monthKeys = userGroup
                    .OrderBy(x => x.year).ThenBy(x => x.month)
                    .Select((entry, index) => new
                    {
                        Key = $"{entry.year:D4}_{entry.month:D2}",
                        Title = $"{GetMonthName(entry.month)} {entry.year}",
                        Column = index + 2 // +1 — задача, +1 — Excel 1-based
                    }).ToList();

                var colMap = monthKeys.ToDictionary(x => x.Key, x => x.Column);
                int totalCol = colMap.Values.Max() + 1;

                // Заголовки
                ws.Cells[1, 1].Value = "Задача";
                foreach (var mk in monthKeys)
                {
                    ws.Cells[1, mk.Column].Value = mk.Title;
                }
                ws.Cells[1, totalCol].Value = "Итого";

                // Данные
                int row = 2;
                foreach (var task in taskData.Keys)
                {
                    double total = 0;
                    var valuesByMonth = new Dictionary<int, double>();

                    // Сначала собираем значения по каждому месяцу
                    foreach (var mk in monthKeys)
                    {
                        string fullKey = $"{mk.Key}_{user}";
                        if (taskData[task].TryGetValue(fullKey, out double hours))
                        {
                            valuesByMonth[mk.Column] = hours;
                            total += hours;
                        }
                    }

                    // Если у пользователя 0 часов по этой задаче — пропускаем
                    if (total == 0) continue;

                    ws.Cells[row, 1].Value = task;

                    foreach (var kvp in valuesByMonth)
                    {
                        ws.Cells[row, kvp.Key].Value = kvp.Value;
                    }

                    ws.Cells[row, totalCol].Value = total;
                    row++;
                }

                ws.Columns.AutoFit();
            }
        }

        private string GetNamedValue(Workbook wb, string name)
        {
            try
            {
                var range = wb.Names.Item(name)?.RefersToRange;
                var val = range?.Value;
                if (val is object[,] arr)
                    return arr[1, 1]?.ToString();
                return val?.ToString();
            }
            catch
            {
                return "";
            }
        }

        private string GetMonthName(int month)
        {
            return System.Globalization.CultureInfo
                .GetCultureInfo("ru-RU")
                .DateTimeFormat.GetMonthName(month);
        }

        private void ProcessFile(string filePath, Application app, Dictionary<string, Dictionary<string, double>> taskData, SortedSet<(int, int, string, string)> columnMeta)
        {
            var wb = app.Workbooks.Open(filePath, ReadOnly: true);
            try
            {
                string user = GetNamedValue(wb, "UserName");
                int year = int.Parse(GetNamedValue(wb, "Year"));
                string monthStr = (GetNamedValue(wb, "Month"));
                int month = DateTime.ParseExact(monthStr, "MMMM", new CultureInfo("ru-RU")).Month;
                string title = $"{GetMonthName(month)} {year} ({user})";
                string userKey = $"{year:D4}_{month:D2}_{user}";

                columnMeta.Add((year, month, user, title));

                var range = wb.Names.Item("ProductsTime")?.RefersToRange;
                if (range == null) return;

                for (int r = 1; r <= range.Rows.Count; r++)
                {
                    string task = Convert.ToString(range.Cells[r, 1].Value);
                    if (string.IsNullOrWhiteSpace(task)) continue;

                    if (!double.TryParse(Convert.ToString(range.Cells[r, 2].Value), out double hours)) continue;

                    if (!taskData.ContainsKey(task))
                        taskData[task] = new Dictionary<string, double>();

                    if (taskData[task].ContainsKey(userKey))
                        taskData[task][userKey] += hours;
                    else
                        taskData[task][userKey] = hours;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обработке файла {filePath}:\n{ex.Message}");
            }
            finally
            {
                wb.Close(false);
            }
        }

        public void LoadProjectsAndProductsFromWorkbook()
        {
            var storage = ProjectProductStorage;
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Name namedRange = null;

            try
            {
                namedRange = workbook.Names.Item("ProjectsAndProducts");
            }
            catch
            {
                MessageBox.Show("Именованный диапазон 'ProjectsAndProducts' не найден.");
                return;
            }

            Range range = namedRange.RefersToRange;
            int rowCount = range.Rows.Count;

            // Подсчёт новых проектов и изделий
            Dictionary<string, HashSet<string>> addedProjects = new Dictionary<string, HashSet<string>>();

            for (int i = 1; i <= rowCount; i++)
            {
                string projectName = Convert.ToString((range.Cells[i, 1] as Range)?.Value2)?.Trim();
                string productName = Convert.ToString((range.Cells[i, 2] as Range)?.Value2)?.Trim();

                if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(productName))
                    continue;

                if (!storage.ProjectsProducts.ContainsKey(projectName))
                {
                    storage.ProjectsProducts[projectName] = new List<string> { productName };
                    addedProjects[projectName] = new HashSet<string> { productName };
                }
                else
                {
                    if (!storage.ProjectsProducts[projectName].Contains(productName))
                    {
                        storage.ProjectsProducts[projectName].Add(productName);

                        if (!addedProjects.ContainsKey(projectName))
                            addedProjects[projectName] = new HashSet<string>();

                        addedProjects[projectName].Add(productName);
                    }
                }
            }

            storage.SaveToFile();

            // Формирование отчета
            if (addedProjects.Count == 0)
            {
                MessageBox.Show("Новые данные не были добавлены.");
                return;
            }

            StringBuilder report = new StringBuilder();
            report.AppendLine($"Успешно добавлено {addedProjects.Count} проект(ов):");

            int index = 1;
            foreach (var kvp in addedProjects.OrderBy(p => p.Key))
            {
                report.AppendLine($"{index}. {kvp.Key} - {kvp.Value.Count} изд.");
                index++;
            }

            int totalProjects = storage.ProjectsProducts.Count;
            int totalProducts = storage.ProjectsProducts.Sum(p => p.Value.Count);

            report.AppendLine();
            report.AppendLine($"Всего проектов: {totalProjects}, всего изделий: {totalProducts}");

            MessageBox.Show(report.ToString(), "Обновление справочника", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void UpdateWorkbookFromStorage()
        {
            var storage = ProjectProductStorage;
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Name namedRange = null;

            try
            {
                namedRange = workbook.Names.Item("ProjectsAndProducts");
            }
            catch
            {
                MessageBox.Show("Именованный диапазон 'ProjectsAndProducts' не найден.");
                return;
            }

            Range range = namedRange.RefersToRange;
            Worksheet sheet = range.Worksheet;

            // Список пар (проект, изделие)
            var rows = new List<(string Project, string Product)>();
            foreach (var kvp in storage.ProjectsProducts)
            {
                string project = kvp.Key;
                foreach (string product in kvp.Value)
                {
                    rows.Add((project, product));
                }
            }

            // Естественная сортировка
            rows = rows
                .OrderBy(r => r.Project, new NaturalStringComparer())
                .ThenBy(r => r.Product, new NaturalStringComparer())
                .ToList();

            int existingRows = range.Rows.Count;
            int requiredRows = rows.Count;
            int startRow = range.Row;
            int startColumn = range.Column;

            if (requiredRows > existingRows)
            {
                int newLastRow = startRow + requiredRows - 1;
                Range newRange = sheet.Range[
                    sheet.Cells[startRow, startColumn],
                    sheet.Cells[newLastRow, startColumn + 1]
                ];

                namedRange.RefersTo = $"='{sheet.Name}'!{newRange.Address[ReferenceStyle: XlReferenceStyle.xlA1]}";
                range = newRange;
            }

            range.ClearContents();

            for (int i = 0; i < rows.Count; i++)
            {
                (range.Cells[i + 1, 1] as Range).Value2 = rows[i].Project;
                (range.Cells[i + 1, 2] as Range).Value2 = rows[i].Product;
            }

            try
            {
                foreach (WorkbookConnection connection in workbook.Connections)
                {
                    connection.Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при обновлении подключений Power Query: " + ex.Message);
            }
        }



        private void buttonUpdateProjectsProductsList_Click(object sender, RibbonControlEventArgs e)
        {
            LoadProjectsAndProductsFromWorkbook();
        }

        public void ReloadStorageFromWorkbookWithConfirmation()
        {
            DialogResult result = MessageBox.Show(
                "Вы хотите удалить старые записи и добавить записи только из этой книги?",
                "Обновление данных",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result != DialogResult.Yes)
                return;

            var storage = ProjectProductStorage;
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Name namedRange = null;

            try
            {
                namedRange = workbook.Names.Item("ProjectsAndProducts");
            }
            catch
            {
                MessageBox.Show("Именованный диапазон 'ProjectsAndProducts' не найден.");
                return;
            }

            Range range = namedRange.RefersToRange;
            int rowCount = range.Rows.Count;

            // Очищаем старые записи
            storage.ProjectsProducts.Clear();

            // Загружаем новые из книги
            for (int i = 1; i <= rowCount; i++)
            {
                string projectName = Convert.ToString((range.Cells[i, 1] as Range)?.Value2)?.Trim();
                string productName = Convert.ToString((range.Cells[i, 2] as Range)?.Value2)?.Trim();

                if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(productName))
                    continue;

                if (!storage.ProjectsProducts.ContainsKey(projectName))
                {
                    storage.ProjectsProducts[projectName] = new List<string> { productName };
                }
                else if (!storage.ProjectsProducts[projectName].Contains(productName))
                {
                    storage.ProjectsProducts[projectName].Add(productName);
                }
            }

            storage.SaveToFile();
            MessageBox.Show("Данные успешно обновлены из книги.", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            UpdateWorkbookFromStorage();
        }

        private void buttonClearAndLoadFromBook_Click(object sender, RibbonControlEventArgs e)
        {
            ReloadStorageFromWorkbookWithConfirmation();
        }

        private void buttonShowCalendar_Click(object sender, RibbonControlEventArgs e)
        {
            var excelApp = Globals.ThisAddIn.Application;
            var activeCell = excelApp.ActiveCell as Range;

            DateTime selectedDate = DateTime.Today;

            if (activeCell != null && activeCell.Value2 != null)
            {
                try
                {
                    if (activeCell.Value2 is double)
                    {
                        selectedDate = DateTime.FromOADate((double)activeCell.Value2);
                    }
                    else if (DateTime.TryParse(activeCell.Value2.ToString(), out DateTime parsedDate))
                    {
                        selectedDate = parsedDate;
                    }
                }
                catch
                {
                    selectedDate = DateTime.Today;
                }
            }

            selectedDate = selectedDate.Date;

            var calendarForm = new MultiMonthCalendarForm(selectedDate);
            var result = calendarForm.ShowDialog(); // модальное окно

            if (result == DialogResult.OK && calendarForm.SelectedDate.HasValue && activeCell != null)
            {
                var onlyDate = calendarForm.SelectedDate.Value.Date;

                activeCell.NumberFormat = "dd.mm.yyyy"; // Устанавливаем формат ячейки как дата
                activeCell.Value = onlyDate;

                // Принудительно пересчитываем лист
                activeCell.Worksheet.Calculate();
            }
        }
    }

}
