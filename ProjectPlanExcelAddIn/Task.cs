using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectPlanExcelAddIn
{
    public class Task
    {
        public int RowIndex { get; set; }
        public string Name { get; set; }
        public int Duration { get; set; }
        public int Priority { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

        public Task(int rowIndex, string name, int duration, int priority)
        {
            RowIndex = rowIndex;
            Name = name;
            Duration = duration;
            Priority = priority;
        }
    }
}
