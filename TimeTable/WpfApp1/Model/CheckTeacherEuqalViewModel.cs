using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WpfApp1.Model.FileMenegers;

namespace WpfApp1.Model
{
    public class CheckTeacherEuqalViewModel
    {
        public TeachersFileMeneger teachersFileMeneger = new TeachersFileMeneger();
        private ObservableCollection<TeachersAll> teachers;
        private ObservableCollection<TeachersAll> badTeachers;
        public ObservableCollection<TeachersAll> Teachers { get { return teachers; } }
        public ObservableCollection<TeachersAll> BadTeachers { get { return badTeachers; } }
         public CheckTeacherEuqalViewModel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            teachersFileMeneger = new TeachersFileMeneger();
            teachers = new ObservableCollection<TeachersAll>();
            badTeachers = new ObservableCollection<TeachersAll>();
            InitIdialTeachersListAsync();
        }
        public async Task InitIdialTeachersListAsync()
        {
            try
            {
                List<string> file = await teachersFileMeneger.Read();
                foreach (var item in file)
                {
                    teachers.Add(new TeachersAll(item));
                }
            }
            catch (Exception ex)
            {

            }

        }
        public void InitBadTeachersList()
        {
            using (ExcelPackage excelPackage = new ExcelPackage(CheckGroupOnEqual.TimetableFile))
            {
                try
                {
                    GetTeachersFromTimetable(excelPackage);
                }
                catch (Exception ex)
                {

                }

            }
        }
        public void GetTeachersFromTimetable(ExcelPackage excelPackage)
        {
            try
            {
                List<TeachersAll> teachersFromTimetable = new List<TeachersAll>();
                int listCount = excelPackage.Workbook.Worksheets.Count();
                List<ExcelWorksheet> anotherWorksheet = new List<ExcelWorksheet>();
                for (int i = 0; i < listCount; i++)
                {
                    anotherWorksheet.Add(excelPackage.Workbook.Worksheets[i]);
                }
                foreach (var item in anotherWorksheet)
                {
                    int col = item.Dimension.End.Column;
                    for (int i = 1; i < col; i++)
                    {
                        double width = item.Column(i).Width;
                        for (int j = 9; j < 87; j=j+2)
                        {
                            if (item.Cells[j, i].Value != null && item.Column(i).Width > 0)
                            {
                                if (!item.Cells[j, i].Value.ToString().ToLower().Contains("понедельник") && 
                                    !item.Cells[j, i].Value.ToString().ToLower().Contains("вторник") && 
                                    !item.Cells[j, i].Value.ToString().ToLower().Contains("среда") && 
                                    !item.Cells[j, i].Value.ToString().ToLower().Contains("четверг") && 
                                    !item.Cells[j, i].Value.ToString().ToLower().Contains("пятница") && 
                                    !item.Cells[j, i].Value.ToString().ToLower().Contains("суббота") &&
                                    !Regex.IsMatch(item.Cells[j, i].Value.ToString(), @"^[0-9]{3,5}.[0-9]{3,5}", RegexOptions.IgnoreCase) && 
                                    item.Cells[j, i].Value.ToString() != "1" && 
                                    item.Cells[j, i].Value.ToString() != "2" && 
                                    item.Cells[j, i].Value.ToString() != "3" && 
                                    item.Cells[j, i].Value.ToString() != "4" && 
                                    item.Cells[j, i].Value.ToString() != "5" && 
                                    item.Cells[j, i].Value.ToString() != "6" && 
                                    item.Cells[j, i].Value.ToString() != "7")
                                {
                                    bool contains = false;
                                    foreach (var teacher in teachersFromTimetable)
                                    {
                                        if (teacher.Names == item.Cells[j, i].Value.ToString())
                                        {
                                            contains = true;
                                        }
                                    }
                                }

                            }
                        }
                        
                    }
                }
                foreach (var item in teachersFromTimetable)
                {
                    bool flag = false;
                    foreach (var teachers in teachers)
                    {
                        if (item.Names == teachers.Names)
                        {
                            flag = true;
                            break;
                        }
                    }
                    if (!flag)
                    {
                        badTeachers.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
