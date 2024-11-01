using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace ProjectPlanExcelAddIn
{
    // Календарь
    public class BusinessCalendar
    {
        private readonly HashSet<DateTime> _holidays;

        public BusinessCalendar()
        {
            // Загружаем праздничные дни с листа "calendar"
            _holidays = LoadHolidaysFromSheet();
        }

        // Метод для сдвига даты на указанное количество рабочих дней
        public DateTime ShiftDate(DateTime startDate, int shiftDays)
        {
            DateTime currentDate = startDate;
            int daysMoved = 0;

            // Если shiftDays равно 0, то мы проверяем только ближайший рабочий день, начиная с текущей даты
            if (shiftDays == 0)
            {
                while (!IsWorkingDay(currentDate))
                {
                    currentDate = currentDate.AddDays(1); // Переходим к следующему дню, пока не найдем рабочий
                }
                return currentDate;
            }

            // Обработка случаев, когда shiftDays не равно 0
            while (daysMoved < Math.Abs(shiftDays))
            {
                currentDate = shiftDays > 0 ? currentDate.AddDays(1) : currentDate.AddDays(-1);

                if (IsWorkingDay(currentDate))
                {
                    daysMoved++;
                }
            }

            return currentDate;
        }

        // Проверка, является ли день рабочим
        private bool IsWorkingDay(DateTime date)
        {
            // День считается рабочим, если он не находится в списке праздников
            return !_holidays.Contains(date.Date);
        }

        // Метод для добавления праздника (если понадобится вручную добавить)
        public void AddHoliday(DateTime holiday)
        {
            _holidays.Add(holiday.Date);
        }

        // Метод для загрузки праздничных дней с листа "calendar"
        private HashSet<DateTime> LoadHolidaysFromSheet()
        {
            HashSet<DateTime> holidays = new HashSet<DateTime>();

            try
            {
                Application excelApp = Globals.ThisAddIn.Application;
                Worksheet calendarSheet = excelApp.Worksheets["calendar"] as Worksheet;

                if (calendarSheet != null)
                {
                    int rowIndex = 1; // Начинаем с первой строки
                    while (true)
                    {
                        Range dateCell = calendarSheet.Cells[rowIndex, 1];

                        if (DateTime.TryParse(dateCell.Text.ToString(), out DateTime holidayDate))
                        {
                            holidays.Add(holidayDate.Date);
                        }
                        else
                        {
                            // Останавливаемся, если достигли пустой ячейки (конец данных)
                            break;
                        }

                        rowIndex++;
                    }
                }
                else
                {
                    throw new Exception("Лист 'calendar' не найден. Убедитесь, что лист существует.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке календаря: {ex.Message}");
            }

            return holidays;
        }
    }
}
