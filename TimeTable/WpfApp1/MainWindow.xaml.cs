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
using System.Windows.Navigation;
using System.Windows.Shapes;
using WpfApp1.Model;
using WpfApp1.Model.ExcelFile;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public FileManager fileManager = new FileManager();
        public GroupFileMeneger groupFileManager = new GroupFileMeneger();
        public FileInfo timetableFile;
        public List<Group> GroupList = new List<Group>();
        public MainWindow()
        {
            InitializeComponent();

        }
        
        private async void addFiles(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlgBin = new Microsoft.Win32.OpenFileDialog();
            dlgBin.FileName = "Document"; // Default file name
            dlgBin.DefaultExt = ".xlsx"; // Default file extension
            dlgBin.Filter = "Excel Files|*.xlsx;"; // Filter files by extension

            // Show save file dialog box
            Nullable<bool> result = dlgBin.ShowDialog();

            if (result == true)
            {
                string path = dlgBin.FileName;
                OldFile newFile = new OldFile(DateTime.Now, path, path);
                bool contains = false;
                var read= fileManager.Read().Result;
                for (int i = 0; i < read.Count(); i++)
                {
                    if (read[i].Path == newFile.Path)
                    {
                        contains = true;
                    }
                }
                if (!contains)
                {
                    fileManager.Save(newFile);
                }
                await OpenFile(path);
            }

        }
        public async Task OpenFile(string path)
        {
           
            //Тут считываем файл
            timetableFile = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage(timetableFile))
            {
                try
                {
                    ExcelWorksheet anotherWorksheet = excelPackage.Workbook.Worksheets[0];
                    //test.Text = anotherWorksheet.Cells[7,30].Value.ToString();
                    GetGroupFromTimetable(excelPackage);
                }
               catch (Exception ex)
                {

                }
 
            }
        }

        public void GetGroupFromTimetable(ExcelPackage excelPackage)
        {
            try
            {
                //roomText.Text = "";
                List<Group> groups = new List<Group>();
                excelPackage.Workbook.Worksheets.Count();
                List<ExcelWorksheet> anotherWorksheet = new List<ExcelWorksheet>();
                for (int i = 0; i < excelPackage.Workbook.Worksheets.Count();i++)
                {
                    anotherWorksheet.Add(excelPackage.Workbook.Worksheets[i]);
                }
                foreach (var item in anotherWorksheet)
                {
                    int col = item.Dimension.End.Column;
                    for (int i = 1; i < col; i++)
                    {
                        double width = item.Column(i).Width;

                        if (item.Cells[7, i].Value != null&& item.Column(i).Width>0)
                        {
                            if (!item.Cells[7, i].Value.ToString().Contains("время")&&!item.Cells[7, i].Value.ToString().Contains("пара")&&item.Cells[7, i].Value.ToString()!="1"&&item.Cells[7, i].Value.ToString() != "2")
                            {
                                bool contains = false;
                                foreach (var group in groups)
                                {
                                    if (group.Name== item.Cells[7, i].Value.ToString())
                                    {
                                        contains = true;
                                    }
                                }
                                if (!contains)
                                {
                                    if (item.Cells[7, i].Value.ToString().Contains("51") || item.Cells[7, i].Value.ToString().Contains("52"))
                                    {
                                        groups.Add(new Group(item.Cells[7, i].Value.ToString(), 5));
                                    }
                                    if (item.Cells[7, i].Value.ToString().Contains("41")|| item.Cells[7, i].Value.ToString().Contains("42"))
                                    {
                                        groups.Add(new Group(item.Cells[7, i].Value.ToString(),4));
                                    }
                                    if (item.Cells[7, i].Value.ToString().Contains("31") || item.Cells[7, i].Value.ToString().Contains("32"))
                                    {
                                        groups.Add(new Group(item.Cells[7, i].Value.ToString(), 3));
                                    }
                                    if (item.Cells[7, i].Value.ToString().Contains("21") || item.Cells[7, i].Value.ToString().Contains("22"))
                                    {
                                        groups.Add(new Group(item.Cells[7, i].Value.ToString(), 2));
                                    }
                                    if (item.Cells[7, i].Value.ToString().Contains("11") || item.Cells[7, i].Value.ToString().Contains("12"))
                                    {
                                        groups.Add(new Group(item.Cells[7, i].Value.ToString(), 1));
                                    }
                                    //roomText.Text = roomText.Text + item.Cells[7, i].Value.ToString() + " ";
                                }
                            }
                            
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        private void CleanList(object sender, RoutedEventArgs e)
        {
            fileManager.Clear();
        }

        private void CommandBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {

        }

        public void ChangeIdialGroupList()
        {

        }
       
        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void AddGroup(object sender, RoutedEventArgs e)
        {
            
        }
        private async void RemoveGroupAsync(object sender, RoutedEventArgs e)
        {
            try
            {
                List<Group> files = await groupFileManager.Read();
            }
            catch (Exception ex)
            {

            }
        }
    }
}
