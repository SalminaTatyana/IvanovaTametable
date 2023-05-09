using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для CheckTimetableOnWindow.xaml
    /// </summary>
    public partial class CheckTimetableOnWindow : Window
    {
        public CheckTimetableOnWindow()
        {
            InitializeComponent();
        }
        public static FileInfo TimetableFile { get; set; }


        public CheckTimetableOnWindow(FileInfo timetableFile)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            TimetableFile = timetableFile;
            InitializeComponent();

        }
    }
}
