using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.Model
{
    public class OldFile
    {
        public DateTime Date { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public OldFile(DateTime date, string name, string path)
        {
            Date = date;
            Name = name;
            Path = path;
        }
    }

   
}
