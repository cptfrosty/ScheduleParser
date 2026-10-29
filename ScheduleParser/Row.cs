using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScheduleParser
{
    internal class Row
    {
        public string? Date { get; set; }
        public string? LessonNumber { get; set; }
        public string? NameLesson { get; set; }
        public string? Teacher { get; set; }
        public string? ClassroomNumber { get; set; }
    }
}
