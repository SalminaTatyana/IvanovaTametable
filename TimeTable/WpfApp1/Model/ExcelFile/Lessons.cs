using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.Model.ExcelFile
{
    public class Lessons
    {
        public string Teacher { get; set; }
        public string Classroom { get; set; }
        public Discipline Discipline { get; set; }
        public Group Group { get; set; }
        public Lessons(string teacher, string classroom, Discipline discipline, Group group)
        {
            Teacher = teacher;
            Classroom = classroom;
            Discipline = discipline;
            Group = group;
        }
    }
}
