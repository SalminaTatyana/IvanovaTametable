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
        public static FileInfo timetableFile;
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
            NowFileName.Text = path;
            timetableFile = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

        private void GoCheckGroupOnEqual(object sender, RoutedEventArgs e)
        {
            if (timetableFile!=null)
            {
                CheckGroupOnEqual view = new CheckGroupOnEqual(timetableFile);
                view.Show();
            }
            
        }
    }
}
