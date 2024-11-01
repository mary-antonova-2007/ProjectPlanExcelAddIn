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
    }
}
