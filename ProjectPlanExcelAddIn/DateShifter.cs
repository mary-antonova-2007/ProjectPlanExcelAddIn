using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectPlanExcelAddIn
{
    // Класс для сдвига дат по рабочему календарю
    public class DateShifter
    {
        private readonly BusinessCalendar _businessCalendar;
        private readonly int _shiftDays;

        // Конструктор принимает рабочий календарь и количество дней для сдвига
        public DateShifter(BusinessCalendar businessCalendar, int shiftDays)
        {
            _businessCalendar = businessCalendar;
            _shiftDays = shiftDays;
        }

        // Метод для обработки выделенного диапазона
        public void ShiftSelectedDates(Range selectedRange)
        {
            foreach (Range cell in selectedRange)
            {
                if (DateTime.TryParse(cell.Value?.ToString(), out DateTime cellDate))
                {
                    cell.Value = _businessCalendar.ShiftDate(cellDate, _shiftDays);
                }
                else
                {
                    // Игнорируем ячейки, где не указана дата
                    Console.WriteLine($"Ячейка {cell.Address} не содержит дату. Пропускаем!");
                }
            }
        }
        public void ShiftDatesInRows(Range selectedRange)
        {
            foreach (Range row in selectedRange.Rows)
            {
                foreach (Range cell in row.Cells)
                {
                    DateTime cellDate = DateTime.MaxValue;
                    // Пропускаем пустые ячейки или ячейки, не содержащие даты
                    if (cell.Value == null || !DateTime.TryParse(cell.Value.ToString(), out cellDate))
                        continue;

                    if (cellDate != DateTime.MinValue)
                    {
                        // Сдвигаем дату с учетом рабочего календаря
                        cell.Value = _businessCalendar.ShiftDate(cellDate, _shiftDays);
                    }
                        
                }
            }
        }
    }
}
