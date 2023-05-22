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
    public class TeachersOnDoubleLessons
    {
        public TeachersAll Classrooms { get; set; }
        public LessonsAll Lessons { get; set; }
        public object Row { get; set; }
        public object Week { get; set; }
        public object Col { get; set; }
        public object Page { get; set; }
        public TeachersOnDoubleLessons(TeachersAll classrooms, LessonsAll lessons, int row, int col, int page, int week)
        {
            Classrooms = classrooms;
            Lessons = lessons;
            Row = row;
            Col = col;
            Page = page;
            Week = week;
        }
        public TeachersOnDoubleLessons(TeachersAll classrooms, LessonsAll lessons, string row, string col, string page, string week)
        {
            Classrooms = classrooms;
            Lessons = lessons;
            Row = row;
            Col = col;
            Page = page;
            Week = week;
        }
    }
    public class CheckTimetableOnLessonsOfTeacherViewModel
    {
        public LessonsFileMeneger lessonsFileMeneger { get; set; }
        public TeachersFileMeneger tecahersFileMeneger { get; set; }
        private ObservableCollection<TeachersOnDoubleLessons> classrooms;
        public TeachersOnDoubleLessons SelectedClassrooms { get; set; }
        private ObservableCollection<TeachersAll> classroomsAtFile;
        private ObservableCollection<LessonsAll> lessonsAtFile;
        public ObservableCollection<TeachersOnDoubleLessons> Classrooms { get { return classrooms; } }
        List<ClassroomsAll> classroomsFromTimetable { get; set; }
        List<LessonsAll> lessonsFromTimetable { get; set; }

        public RelayCommand HighlightClassroomsChange { get; set; }
        public List<TeachersOnDoubleLessons> checkClassrooms { get; set; }
        public CheckTimetableOnLessonsOfTeacherViewModel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            lessonsFileMeneger = new LessonsFileMeneger();
            tecahersFileMeneger = new TeachersFileMeneger();
            classrooms = new ObservableCollection<TeachersOnDoubleLessons>();
            classroomsAtFile = new ObservableCollection<TeachersAll>();
            lessonsAtFile = new ObservableCollection<LessonsAll>();
            classroomsFromTimetable = new List<ClassroomsAll>();
            lessonsFromTimetable = new List<LessonsAll>();
            checkClassrooms = new List<TeachersOnDoubleLessons>();
            HighlightClassroomsChange = new RelayCommand(o => HighlightLessons(SelectedClassrooms));
            InitIdialClassroomsListAsync();
        }
        public async Task InitIdialClassroomsListAsync()
        {
            try
            {
                List<string> file = await tecahersFileMeneger.Read();
                foreach (var item in file)
                {
                    classroomsAtFile.Add(new TeachersAll(item));
                }
                List<string> lessons = await lessonsFileMeneger.Read();
                foreach (var item in lessons)
                {
                    lessonsAtFile.Add(new LessonsAll(item));
                }
                IdialClassroomsListAsync();
            }
            catch (Exception ex)
            {

            }

        }
        public async Task IdialClassroomsListAsync()
        {
            try
            {
                classrooms.Add(new TeachersOnDoubleLessons(new TeachersAll("Преподаватель"), new LessonsAll("Занятие"), "Строка", "Столбец", "Страница", "Неделя"));
                using (ExcelPackage excelPackage = new ExcelPackage(CheckTimetableOnLessonsOfTeacher.TimetableFile))
                {
                    int listCount = excelPackage.Workbook.Worksheets.Count();
                    List<ExcelWorksheet> anotherWorksheet = new List<ExcelWorksheet>();
                    for (int i = 0; i < listCount; i++)
                    {
                        anotherWorksheet.Add(excelPackage.Workbook.Worksheets[i]);
                    }
                    foreach (var item in anotherWorksheet)
                    {
                        int countWeek = 1;
                        int col = item.Dimension.End.Column;
                        for (int i = 1; i < col; i++)
                        {

                            double width = item.Column(i).Width;
                            if (item.Cells[7, i].Value != null)
                            {
                                if (
                                       item.Cells[7, i].Value.ToString() == "1")
                                {
                                    countWeek = 1;
                                }
                                if (
                                       item.Cells[7, i].Value.ToString() == "2")
                                {
                                    countWeek = 2;
                                }
                            }

                            for (int j = 8; j < 87; j = j + 2)
                            {

                                if (item.Cells[j + 1, i].Value != null && item.Cells[j, i].Value != null)
                                {

                                    if (!item.Cells[j + 1, i].Value.ToString().ToLower().Contains("понедельник") &&
                                       !item.Cells[j + 1, i].Value.ToString().ToLower().Contains("вторник") &&
                                       !item.Cells[j + 1, i].Value.ToString().ToLower().Contains("среда") &&
                                       !item.Cells[j + 1, i].Value.ToString().ToLower().Contains("четверг") &&
                                       !item.Cells[j + 1, i].Value.ToString().ToLower().Contains("пятница") &&
                                       !item.Cells[j + 1, i].Value.ToString().ToLower().Contains("суббота") &&
                                       !Regex.IsMatch(item.Cells[j + 1, i].Value.ToString(), @"^[0-9]{3,5}.[0-9]{3,5}", RegexOptions.IgnoreCase) &&
                                       item.Cells[j + 1, i].Value.ToString() != "1" &&
                                       item.Cells[j + 1, i].Value.ToString() != "2" &&
                                       item.Cells[j + 1, i].Value.ToString() != "3" &&
                                       item.Cells[j + 1, i].Value.ToString() != "4" &&
                                       item.Cells[j + 1, i].Value.ToString() != "5" &&
                                       item.Cells[j + 1, i].Value.ToString() != "6" &&
                                       item.Cells[j + 1, i].Value.ToString() != "7")
                                    {
                                        Regex regex = new Regex(@"[1-6].[0-9]{1,3}");
                                        Regex regexSport = new Regex(@"[1-6].[0-9]{1,3}");
                                        MatchCollection matches = regex.Matches(item.Cells[j + 1, i].Value.ToString());
                                        MatchCollection matchesSport = regexSport.Matches(item.Cells[j + 1, i].Value.ToString());
                                        string str = "";
                                        if (matches.Count > 0)
                                        {
                                            int index = item.Cells[j + 1, i].Value.ToString().IndexOf(matches[0].Value[0]);
                                            str = item.Cells[j + 1, i].Value.ToString().Substring(0,index).Trim();
                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        else if (item.Cells[j + 1, i].Value.ToString().ToLower().Contains("дист"))
                                        {
                                            int index = item.Cells[j + 1, i].Value.ToString().IndexOf("дист");
                                            str = item.Cells[j + 1, i].Value.ToString().Substring(0,index).Trim();
                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        else if (item.Cells[j + 1, i].Value.ToString().ToLower().Contains("зал"))
                                        {
                                            int index = item.Cells[j + 1, i].Value.ToString().IndexOf("зал");
                                            str = item.Cells[j + 1, i].Value.ToString().Substring(0,index).Trim();
                                            if (item.Cells[j + 1, i].Value.ToString().Substring(0, index).Trim().Contains("1-") || item.Cells[j + 1, i].Value.ToString().Substring(0, index).Trim().Contains("3-"))
                                            {
                                                str = item.Cells[j + 1, i].Value.ToString().Substring(0, index ).Trim();
                                            }

                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        else if (item.Cells[j + 1, i].Value.ToString().ToLower().Contains("базовая кафедра"))
                                        {
                                            int index = item.Cells[j + 1, i].Value.ToString().IndexOf("базовая кафедра");
                                            str = item.Cells[j + 1, i].Value.ToString().Substring(0,index).Trim();
                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        else 
                                        {
                                            str = item.Cells[j + 1, i].Value.ToString().Trim();
                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        Regex regex1 = new Regex(@"-л.{0,2}$|-п.{0,2}$");
                                        MatchCollection matches1 = regex1.Matches(item.Cells[j, i].Value.ToString());
                                        if (matches1.Count > 0)
                                        {
                                            int index = item.Cells[j, i].Value.ToString().IndexOf(matches1[0].Value);
                                            string str1 = item.Cells[j, i].Value.ToString().Substring(0, index).Trim();
                                            lessonsFromTimetable.Add(new LessonsAll(str1));
                                            checkClassrooms.Add(new TeachersOnDoubleLessons(new TeachersAll(String.IsNullOrEmpty(str) ? "нет преподавателя" : str), new LessonsAll(str1), j, i, item.Index + 1, countWeek));
                                        }
                                    }


                                }


                            }

                        }
                    }

                    foreach (var item1 in checkClassrooms)
                    {
                        foreach (var item2 in checkClassrooms)
                        {
                            if (item1 != item2)
                            {

                                if (item1.Classrooms.Names == item2.Classrooms.Names)
                                {
                                    if (item1.Lessons.Names != item2.Lessons.Names)
                                    {
                                        if ((int)item1.Week == (int)item2.Week)
                                        {
                                            if ((int)item1.Row == (int)item2.Row)
                                            {
                                                if (!classrooms.Contains(item1) && !classrooms.Contains(item2))
                                                {
                                                    classrooms.Add(item1);
                                                    classrooms.Add(item2);

                                                }
                                                break;
                                            }
                                        }

                                    }
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
        public void HighlightLessons(TeachersOnDoubleLessons SelectedClassrooms)
        {
            try
            {
                if (SelectedClassrooms != null)
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(CheckTimetableOnLessonsOfTeacher.TimetableFile))
                    {
                        int listCount = excelPackage.Workbook.Worksheets.Count();
                        List<ExcelWorksheet> anotherWorksheet = new List<ExcelWorksheet>();
                        for (int i = 0; i < listCount; i++)
                        {
                            anotherWorksheet.Add(excelPackage.Workbook.Worksheets[i]);
                        }

                        if (anotherWorksheet[(int)SelectedClassrooms.Page - 1].Cells[(int)SelectedClassrooms.Row, (int)SelectedClassrooms.Col].Value != null)
                        {

                            anotherWorksheet[(int)SelectedClassrooms.Page - 1].Cells[(int)SelectedClassrooms.Row, (int)SelectedClassrooms.Col].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFD4948"));
                        }
                        excelPackage.SaveAs(CheckTimetableOnLessonsOfTeacher.TimetableFile);
                        excelPackage.Dispose();
                    }
                }

            }
            catch (Exception ex)
            {

            }

        }
    }
}
