using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.Model.ExcelFile
{
    public class Classrooms
    {
        public string Names { get; set; }
        public int PeopleNumber { get; set; }
        public bool Practics { get; set; }
        public bool Labs { get; set; }


        public Classrooms(string name, bool practics, bool labs, int peopleNumber)
        {
            Names = name;
            Practics = practics;
            Labs = labs;
            PeopleNumber = peopleNumber;
        }
    }
}
