using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace ProjectPlanExcelAddIn
{
    public class TaskScheduler
    {
        private readonly Application _excelApp;
        private BusinessCalendar _calendar; // Локальный экземпляр календаря

        public TaskScheduler(Application excelApp)
        {
            _excelApp = excelApp;
        }

        public void ScheduleTasks(Worksheet worksheet)
        {
            // Инициализируем календарь, если он еще не создан
            if (_calendar == null)
            {
                _calendar = new BusinessCalendar();
            }

            int rowIndex = 7;
            List<Task> currentGroupTasks = new List<Task>();
            DateTime? groupStartDate = null;

            while (true)
            {
                Range taskNameCell = worksheet.Cells[rowIndex, 4];
                Range daysCell = worksheet.Cells[rowIndex, 5];
                Range priorityCell = worksheet.Cells[rowIndex, 6];
                Range startDateCell = worksheet.Cells[rowIndex, 7];
                Range endDateCell = worksheet.Cells[rowIndex, 8];

                // Проверяем, является ли строка пустой (отсутствует наименование задачи)
                if (string.IsNullOrWhiteSpace(taskNameCell.Text.ToString()))
                {
                    // Если у нас есть собранные задачи группы, планируем их
                    if (currentGroupTasks.Count > 0)
                    {
                        PlanGroupTasks(currentGroupTasks, groupStartDate ?? DateTime.Today.AddDays(1));
                        currentGroupTasks.Clear();
                    }

                    // Проверяем наличие даты начала планирования для новой группы
                    if (DateTime.TryParse(startDateCell.Text.ToString(), out DateTime parsedDate))
                    {
                        groupStartDate = parsedDate;
                    }
                    else
                    {
                        groupStartDate = DateTime.Today.AddDays(1);
                        startDateCell.Value = groupStartDate;
                    }
                }
                else
                {
                    int priority = 1;
                    // Собираем данные задачи
                    if (int.TryParse(daysCell.Text.ToString(), out int duration) &&
                        int.TryParse(priorityCell.Text.ToString(), out priority))
                    {
                        var task = new Task(rowIndex, taskNameCell.Text.ToString(), duration, priority);
                        currentGroupTasks.Add(task);
                    }
                }

                rowIndex++;

                // Проверяем, не достигли ли конца данных
                if (string.IsNullOrWhiteSpace(taskNameCell.Text.ToString()) && string.IsNullOrWhiteSpace(worksheet.Cells[rowIndex, 4].Text.ToString()))
                {
                    // Планируем последнюю группу, если она не была обработана
                    if (currentGroupTasks.Count > 0)
                    {
                        PlanGroupTasks(currentGroupTasks, groupStartDate ?? DateTime.Today.AddDays(1));
                    }
                    break;
                }
            }
        }

        private void PlanGroupTasks(List<Task> tasks, DateTime startDate)
        {
            // Сортируем задачи по приоритету (чем меньше, тем срочнее)
            tasks.Sort((a, b) => a.Priority.CompareTo(b.Priority));

            DateTime currentDate = startDate;
            foreach (var task in tasks)
            {
                // Находим ближайший рабочий день для начала задачи, если текущая дата - выходной
                currentDate = _calendar.ShiftDate(currentDate, 0);

                // Устанавливаем дату начала задачи
                task.StartDate = currentDate;

                // Рассчитываем дату окончания задачи с учетом количества дней на выполнение
                task.EndDate = _calendar.ShiftDate(task.StartDate, task.Duration - 1);

                // Обновляем ячейки в Excel
                var worksheet = _excelApp.ActiveSheet as Worksheet;
                if (worksheet != null)
                {
                    worksheet.Cells[task.RowIndex, 7].Value = task.StartDate;
                    worksheet.Cells[task.RowIndex, 8].Value = task.EndDate;
                }

                // Обновляем текущую дату для следующей задачи
                currentDate = task.EndDate.AddDays(1);
                // Переносим текущую дату на следующий рабочий день, если она попадает на выходной
                currentDate = _calendar.ShiftDate(currentDate, 0);
            }
        }
    }
}
