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
            int maxIndex = worksheet.UsedRange.Rows.Count;
            if (_calendar == null)
            {
                _calendar = new BusinessCalendar();
            }

            int rowIndex = 7;
            List<Task> currentGroupTasks = new List<Task>();
            List<(DateTime Start, DateTime End)> occupiedDates = new List<(DateTime, DateTime)>();
            DateTime? groupStartDate = null;

            while (true)
            {
                Range taskNameCell = worksheet.Cells[rowIndex, 4];
                Range daysCell = worksheet.Cells[rowIndex, 5];
                Range priorityCell = worksheet.Cells[rowIndex, 6];
                Range startDateCell = worksheet.Cells[rowIndex, 7];
                Range endDateCell = worksheet.Cells[rowIndex, 8];
                Range groupNameCell = worksheet.Cells[rowIndex, 1]; // Столбец A - название группы

                // Если строка пустая
                if (string.IsNullOrWhiteSpace(taskNameCell.Text.ToString()))
                {
                    // Если текущая группа задач не пуста, планируем ее
                    if (currentGroupTasks.Count > 0)
                    {
                        PlanGroupTasks(currentGroupTasks, groupStartDate ?? DateTime.Today.AddDays(1), occupiedDates);
                        currentGroupTasks.Clear();
                        occupiedDates.Clear();
                    }

                    // Пропускаем пустые строки, пока не найдем строку с названием группы
                    while (string.IsNullOrWhiteSpace(groupNameCell.Text.ToString()))
                    {
                        rowIndex++;
                        groupNameCell = worksheet.Cells[rowIndex, 1];
                        if (rowIndex > maxIndex) break;
                    }

                    // Если мы нашли строку с названием группы, то устанавливаем дату начала для этой группы
                    if (rowIndex <= maxIndex && !string.IsNullOrWhiteSpace(groupNameCell.Text.ToString()))
                    {
                        // Проверяем, есть ли уже дата начала в строке с названием группы
                        if (string.IsNullOrWhiteSpace(startDateCell.Text.ToString()))
                        {
                            groupStartDate = DateTime.Today.AddDays(1); // Завтрашний день
                            startDateCell.Value = groupStartDate; // Заполняем дату начала только в строку с названием группы
                        }
                        else
                        {
                            if (DateTime.TryParse(startDateCell.Text.ToString(), out DateTime parsedDate))
                            {
                                groupStartDate = parsedDate;
                            }
                        }
                    }
                }
                else
                {
                    // Проверяем, заполнены ли даты для задачи
                    if (DateTime.TryParse(startDateCell.Text.ToString(), out DateTime existingStart))
                    {
                        // Инициализируем existingEnd значением по умолчанию
                        DateTime existingEnd = existingStart;

                        if (DateTime.TryParse(endDateCell.Text.ToString(), out existingEnd))
                        {
                            // Если обе даты заполнены, добавляем их в список занятых и пропускаем задачу
                            occupiedDates.Add((existingStart, existingEnd));
                            Console.WriteLine($"Задача уже запланирована: {taskNameCell.Text}, {existingStart} - {existingEnd}");
                        }
                        else
                        {
                            // Если дата окончания не задана, добавляем только дату начала
                            occupiedDates.Add((existingStart, existingStart));
                            Console.WriteLine($"Задача запланирована только с датой начала: {taskNameCell.Text}, {existingStart}");
                        }
                    }
                    else
                    {
                        // Если даты не заполнены, собираем данные задачи
                        if (int.TryParse(daysCell.Text.ToString(), out int duration))
                        {
                            // Инициализируем priority значением по умолчанию
                            int priority = 1;

                            if (int.TryParse(priorityCell.Text.ToString(), out priority))
                            {
                                currentGroupTasks.Add(new Task(rowIndex, taskNameCell.Text.ToString(), duration, priority));
                                Console.WriteLine($"Добавлена задача для планирования: {taskNameCell.Text}, Дней: {duration}, Приоритет: {priority}");
                            }
                            else
                            {
                                Console.WriteLine($"Приоритет задачи {taskNameCell.Text} не задан, используется значение по умолчанию: {priority}");
                                currentGroupTasks.Add(new Task(rowIndex, taskNameCell.Text.ToString(), duration, priority));
                            }
                        }
                    }
                }

                rowIndex++;

                // Проверяем конец данных
                if (string.IsNullOrWhiteSpace(taskNameCell.Text.ToString()) &&
                    string.IsNullOrWhiteSpace(worksheet.Cells[rowIndex, 4].Text.ToString()))
                {
                    // Планируем оставшиеся задачи
                    if (currentGroupTasks.Count > 0)
                    {
                        PlanGroupTasks(currentGroupTasks, groupStartDate ?? DateTime.Today.AddDays(1), occupiedDates);
                    }
                    break;
                }

                if (rowIndex > maxIndex) break;
            }
        }



        private void PlanGroupTasks(List<Task> tasks, DateTime startDate, List<(DateTime Start, DateTime End)> occupiedDates)
        {
            tasks.Sort((a, b) => a.Priority.CompareTo(b.Priority));
            DateTime currentDate = startDate;

            foreach (var task in tasks)
            {
                while (!IsDateAvailable(currentDate, task.Duration, occupiedDates))
                {
                    currentDate = _calendar.ShiftDate(currentDate.AddDays(1), 0);
                    Console.WriteLine($"Перенос на следующий день: {currentDate}");
                }

                task.StartDate = currentDate;
                task.EndDate = _calendar.ShiftDate(task.StartDate, task.Duration - 1);

                var worksheet = _excelApp.ActiveSheet as Worksheet;
                if (worksheet != null)
                {
                    worksheet.Cells[task.RowIndex, 7].Value = task.StartDate;
                    worksheet.Cells[task.RowIndex, 8].Value = task.EndDate;
                }

                Console.WriteLine($"Запланирована задача: {task.Name}, Начало: {task.StartDate}, Конец: {task.EndDate}");
                occupiedDates.Add((task.StartDate, task.EndDate));
                currentDate = task.EndDate.AddDays(1);
                currentDate = _calendar.ShiftDate(currentDate, 0);
            }
        }

        private bool IsDateAvailable(DateTime startDate, int duration, List<(DateTime Start, DateTime End)> occupiedDates)
        {
            DateTime endDate = _calendar.ShiftDate(startDate, duration - 1);

            foreach (var (start, end) in occupiedDates)
            {
                if (startDate <= end && endDate >= start)
                {
                    return false;
                }
            }
            return true;
        }
    }
}
