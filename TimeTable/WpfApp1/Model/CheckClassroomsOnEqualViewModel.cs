using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WpfApp1.Model.ExcelFile;
using WpfApp1.Model.FileMenegers;

namespace WpfApp1.Model
{
    public class CheckClassroomsOnEqualViewModel
    {
        public ClassroomsFileMeneger classroomsFileMeneger = new ClassroomsFileMeneger();
        private ObservableCollection<ClassroomsAll> classrooms;
        private ObservableCollection<ClassroomsAll> badLessons;
        public ObservableCollection<ClassroomsAll> Classrooms { get { return classrooms; } }
        public ObservableCollection<ClassroomsAll> BadClassrooms { get { return badLessons; } }
        public CheckClassroomsOnEqualViewModel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            classroomsFileMeneger = new ClassroomsFileMeneger();
            classrooms = new ObservableCollection<ClassroomsAll>();
            badLessons = new ObservableCollection<ClassroomsAll>();
            InitIdialClassroomsListAsync();

        }
        public async Task InitIdialClassroomsListAsync()
        {
            try
            {
                List<Classrooms> file = await classroomsFileMeneger.Read();
                foreach (var item in file)
                {
                    classrooms.Add(new ClassroomsAll(item.Names, (item.Practics ? "пр" : ""), item.Labs ? "лб" : "", item.PeopleNumber));
                }
                InitBadClassroomsList();
            }
            catch (Exception ex)
            {

            }

        }
        public void InitBadClassroomsList()
        {
            using (ExcelPackage excelPackage = new ExcelPackage(CheckClassroomOnEqual.TimetableFile))
            {
                try
                {
                    GetClassroomsFromTimetable(excelPackage);
                }
                catch (Exception ex)
                {

                }

            }
        }
        public void GetClassroomsFromTimetable(ExcelPackage excelPackage)
        {
            try
            {
                List<ClassroomsAll> lessonsFromTimetable = new List<ClassroomsAll>();
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
                        for (int j = 9; j < 88; j = j + 2)
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
                                    Regex regex = new Regex(@"[1-6].[0-9]{1,3}");
                                    Regex regexSport = new Regex(@"[1-6].[0-9]{1,3}");
                                    MatchCollection matches = regex.Matches(item.Cells[j, i].Value.ToString());
                                    MatchCollection matchesSport = regexSport.Matches(item.Cells[j, i].Value.ToString());
                                    if (matches.Count > 0)
                                    {
                                        int index = item.Cells[j, i].Value.ToString().IndexOf(matches[0].Value[0]);
                                        string str = item.Cells[j, i].Value.ToString().Substring(index).Trim();
                                        lessonsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                    }
                                    else if (item.Cells[j, i].Value.ToString().ToLower().Contains("дист"))
                                    {
                                        int index = item.Cells[j, i].Value.ToString().IndexOf("дист");
                                        string str = item.Cells[j, i].Value.ToString().Substring(index).Trim();
                                        lessonsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                    }
                                    else if (item.Cells[j, i].Value.ToString().ToLower().Contains("зал"))
                                    {
                                        int index = item.Cells[j, i].Value.ToString().IndexOf("зал");
                                        string str = item.Cells[j, i].Value.ToString().Substring(index).Trim();
                                        if (item.Cells[j, i].Value.ToString().Substring(0, index).Trim().Contains("1-") || item.Cells[j, i].Value.ToString().Substring(0, index).Trim().Contains("3-"))
                                        {
                                            str = item.Cells[j, i].Value.ToString().Substring(index - 4).Trim();
                                        }

                                        lessonsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                    }
                                    else if (item.Cells[j, i].Value.ToString().ToLower().Contains("базовая кафедра"))
                                    {
                                        int index = item.Cells[j, i].Value.ToString().IndexOf("базовая кафедра");
                                        string str = item.Cells[j, i].Value.ToString().Substring(index).Trim();
                                        lessonsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                    }


                                }

                            }
                        }

                    }
                }
                foreach (var item in lessonsFromTimetable)
                {
                    bool flag = false;
                    foreach (var lesson in classrooms)
                    {
                        if (item.Names == lesson.Names)
                        {
                            flag = true;
                            break;
                        }
                    }
                    if (!flag)
                    {
                        bool flag2 = false;
                        foreach (var bad in badLessons)
                        {
                            if (item.Names == bad.Names)
                            {
                                flag2 = true;
                                break;
                            }
                        }
                        if (!flag2)
                        {
                            badLessons.Add(new ClassroomsAll(item.Names, item.Practics, item.Labs, item.PeopleNumber));
                        }

                    }
                }

            }
            catch (Exception ex)
            {

            }
        }
    }
}
