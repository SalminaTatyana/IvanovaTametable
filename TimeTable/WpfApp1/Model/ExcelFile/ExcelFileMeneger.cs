using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.Model.ExcelFile
{
    public class ExcelFileMeneger
    {
        public int pagesNumber { get; set; }
        public List<ExcelPageContent> excelPagesContent { get; set; }
    }
    public class ExcelPageContent
    {
        public string Name { get; set; }
        public int ColumNumber { get; set; }
        public int PageNumber { get; set; }
        public int RowNumber { get; set; }
        public ExcelPageContent(string name, int columNumber, int pageNumber, int rowNumber)
        {
            Name = name;
            ColumNumber = columNumber;
            PageNumber = pageNumber;
            RowNumber = rowNumber;
        }
    }
}
