using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.Model.ExcelFile
{
    public enum Start : int
    {
        First=1,
        Second=2,
        Third=3,
        Founth=4,
        Fifth=5,
        Sixth=6,
        Seventh =7
    }
    public class Discipline
    {
        public Start Number { get; set; }
        public DateTime Start { get;}
        public DateTime End { get; }
        public Discipline(int number)
        {
            Number = (Start)number;
            switch (Number)
            {
                case ExcelFile.Start.First:
                    Start = DateTime.Parse("6/22/2009 08:00:00"); 
                    End = DateTime.Parse("6/22/2009 09:35:00"); 
                    break;
                case ExcelFile.Start.Second:
                    Start = DateTime.Parse("6/22/2009 09:45:00");
                    End = DateTime.Parse("6/22/2009 11:20:00");
                    break;
                case ExcelFile.Start.Third:
                    Start = DateTime.Parse("6/22/2009 12:00:00");
                    End = DateTime.Parse("6/22/2009 13:35:00");
                    break;
                case ExcelFile.Start.Founth:
                    Start = DateTime.Parse("6/22/2009 13:45:00");
                    End = DateTime.Parse("6/22/2009 15:20:00");
                    break;
                case ExcelFile.Start.Fifth:
                    Start = DateTime.Parse("6/22/2009 15:30:00");
                    End = DateTime.Parse("6/22/2009 17:05:00");
                    break;
                case ExcelFile.Start.Sixth:
                    Start = DateTime.Parse("6/22/2009 17:15:00");
                    End = DateTime.Parse("6/22/2009 18:50:00");
                    break;
                case ExcelFile.Start.Seventh:
                    Start = DateTime.Parse("6/22/2009 19:00:00");
                    End = DateTime.Parse("6/22/2009 20:35:00");
                    break;
                default:
                    break;
            }
        }
    }

}
