using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.Model.ExcelFile
{
    public class Group
    {
        public string Name { get; set; }
        public int Cource { get; set; }
        public int StudentNumber { get; set; }
        public Group(string name, int cource)
        {
            Name = name;
            Cource = cource;
        }
        public Group(string str)
        {
            Name = str;  
        }
    }
}
