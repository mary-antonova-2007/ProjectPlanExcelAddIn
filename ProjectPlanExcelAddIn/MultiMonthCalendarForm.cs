using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ProjectPlanExcelAddIn
{
    public partial class MultiMonthCalendarForm : Form
    {
        private DateTime _baseDate;
        public DateTime? SelectedDate { get; private set; }

        public MultiMonthCalendarForm(DateTime startDate)
        {
            _baseDate = startDate;
            InitializeForm();
        }

        private void InitializeForm()
        {
            this.Text = "Выбор даты";
            this.Size = new System.Drawing.Size(750, 250);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;

            for (int i = 0; i < 3; i++)
            {
                var calendar = new MonthCalendar
                {
                    Location = new System.Drawing.Point(10 + i * 240, 10),
                    MaxSelectionCount = 1,
                    ShowTodayCircle = true,
                    TodayDate = _baseDate
                };

                DateTime displayDate = _baseDate.AddMonths(i);
                calendar.SetDate(displayDate);

                calendar.DateSelected += Calendar_DateSelected;

                this.Controls.Add(calendar);
            }
        }

        private void Calendar_DateSelected(object sender, DateRangeEventArgs e)
        {
            SelectedDate = e.Start;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
