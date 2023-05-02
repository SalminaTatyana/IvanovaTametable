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
    public class ClassroomsOnLessonsType {
        public ClassroomsAll Classrooms { get; set; }
        public LessonsType LessonsType { get; set; }
        public object Row { get; set; }
        public object Col { get; set; }
        public object Page { get; set; }
        public ClassroomsOnLessonsType(ClassroomsAll classrooms, LessonsType lessonsType, int row, int col, int page)
        {
            Classrooms = classrooms;
            LessonsType = lessonsType;
            Row = row;
            Col = col;
            Page = page;
        }
        public ClassroomsOnLessonsType(ClassroomsAll classrooms, LessonsType lessonsType, string row, string col, string page)
        {
            Classrooms = classrooms;
            LessonsType = lessonsType;
            Row = row;
            Col = col;
            Page = page;
        }
    }
    public class CheckClasscoomsOnLessonsTypeViewModel
    {
        public ClassroomsFileMeneger classroomsFileMeneger { get; set; }
        private ObservableCollection<ClassroomsOnLessonsType> classrooms;
        public ClassroomsOnLessonsType SelectedClassrooms { get; set; }
        private ObservableCollection<ClassroomsAll> classroomsAtFile;
        public ObservableCollection<ClassroomsOnLessonsType> Classrooms { get { return classrooms; } }
        List<ClassroomsAll> classroomsFromTimetable { get; set; }
        List<LessonsType> lessonsTypeFromTimetable { get; set; }

        public RelayCommand HighlightClassroomsChange { get; set; }
        public List<ClassroomsOnLessonsType> checkClassrooms { get; set; }
        public CheckClasscoomsOnLessonsTypeViewModel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            classroomsFileMeneger = new ClassroomsFileMeneger();
            classrooms = new ObservableCollection<ClassroomsOnLessonsType>();
            classroomsAtFile = new ObservableCollection<ClassroomsAll>();
            classroomsFromTimetable = new List<ClassroomsAll>();
            lessonsTypeFromTimetable = new List<LessonsType>();
            checkClassrooms = new List<ClassroomsOnLessonsType>();
            HighlightClassroomsChange = new RelayCommand(o => HighlightLessons(SelectedClassrooms));
            InitIdialClassroomsListAsync();
        }
        public async Task InitIdialClassroomsListAsync()
        {
            try
            {
                List<Classrooms> file = await classroomsFileMeneger.Read();
                foreach (var item in file)
                {
                    classroomsAtFile.Add(new ClassroomsAll(item.Names, (item.Practics ? "пр" : ""), item.Labs ? "лб" : "", item.PeopleNumber));
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
                checkClassrooms.Add(new ClassroomsOnLessonsType(new ClassroomsAll("Аудитория", "", "", 0), new LessonsType("Тип занятия"), "Строка", "Столбец", "Страница"));
                using (ExcelPackage excelPackage = new ExcelPackage(CheckClassroomsOnLessonsType.TimetableFile))
                {
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
                            for (int j = 8; j < 87; j=j+2)
                            {
                                if (item.Cells[j+1, i].Value!=null&& item.Cells[j, i].Value != null) { 
                                     if (!item.Cells[j+1, i].Value.ToString().ToLower().Contains("понедельник") &&
                                        !item.Cells[j+1, i].Value.ToString().ToLower().Contains("вторник") &&
                                        !item.Cells[j+1, i].Value.ToString().ToLower().Contains("среда") &&
                                        !item.Cells[j+1, i].Value.ToString().ToLower().Contains("четверг") &&
                                        !item.Cells[j+1, i].Value.ToString().ToLower().Contains("пятница") &&
                                        !item.Cells[j+1, i].Value.ToString().ToLower().Contains("суббота") &&
                                        !Regex.IsMatch(item.Cells[j+1, i].Value.ToString(), @"^[0-9]{3,5}.[0-9]{3,5}", RegexOptions.IgnoreCase) &&
                                        item.Cells[j+1, i].Value.ToString() != "1" &&
                                        item.Cells[j+1, i].Value.ToString() != "2" &&
                                        item.Cells[j+1, i].Value.ToString() != "3" &&
                                        item.Cells[j+1, i].Value.ToString() != "4" &&
                                        item.Cells[j+1, i].Value.ToString() != "5" &&
                                        item.Cells[j+1, i].Value.ToString() != "6" &&
                                        item.Cells[j+1, i].Value.ToString() != "7")
                                    {
                                        Regex regex = new Regex(@"[1-6].[0-9]{1,3}");
                                        Regex regexSport = new Regex(@"[1-6].[0-9]{1,3}");
                                        MatchCollection matches = regex.Matches(item.Cells[j + 1, i].Value.ToString());
                                        MatchCollection matchesSport = regexSport.Matches(item.Cells[j + 1, i].Value.ToString());
                                        string str = "";
                                        if (matches.Count > 0)
                                        {
                                            int index = item.Cells[j + 1, i].Value.ToString().IndexOf(matches[0].Value[0]);
                                            str = item.Cells[j + 1, i].Value.ToString().Substring(index).Trim();
                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        else if (item.Cells[j + 1, i].Value.ToString().ToLower().Contains("дист"))
                                        {
                                            int index = item.Cells[j + 1, i].Value.ToString().IndexOf("дист");
                                            str = item.Cells[j + 1, i].Value.ToString().Substring(index).Trim();
                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        else if (item.Cells[j + 1, i].Value.ToString().ToLower().Contains("зал"))
                                        {
                                            int index = item.Cells[j + 1, i].Value.ToString().IndexOf("зал");
                                            str = item.Cells[j + 1, i].Value.ToString().Substring(index).Trim();
                                            if (item.Cells[j + 1, i].Value.ToString().Substring(0, index).Trim().Contains("1-") || item.Cells[j + 1, i].Value.ToString().Substring(0, index).Trim().Contains("3-"))
                                            {
                                                str = item.Cells[j + 1, i].Value.ToString().Substring(index - 4).Trim();
                                            }

                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        else if (item.Cells[j + 1, i].Value.ToString().ToLower().Contains("базовая кафедра"))
                                        {
                                            int index = item.Cells[j + 1, i].Value.ToString().IndexOf("базовая кафедра");
                                            str = item.Cells[j + 1, i].Value.ToString().Substring(index).Trim();
                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        Regex regex1 = new Regex(@"-лб.{0,2}$|-п.{0,2}$");
                                        MatchCollection matches1 = regex1.Matches(item.Cells[j, i].Value.ToString());
                                        if (matches1.Count > 0)
                                        {
                                            int index = item.Cells[j, i].Value.ToString().IndexOf(matches1[0].Value);
                                            string str1 = item.Cells[j, i].Value.ToString().Substring(index).Trim();
                                            lessonsTypeFromTimetable.Add(new LessonsType(str1));
                                            checkClassrooms.Add(new ClassroomsOnLessonsType(new ClassroomsAll(String.IsNullOrEmpty(str)?"нет аудитории":str, "", "", 0), new LessonsType(str1), j, i, item.Index+1));
                                        }
                                    }

                                }
                                    
                                
                            }

                        }
                    }
                    foreach (var item in checkClassrooms)
                    {
                        bool flag = false;
                        foreach (var classes in classroomsAtFile)
                        {
                            if (item.Classrooms.Names==classes.Names)
                                
                            {
                                if ((item.LessonsType.Names == classes.Labs ||
                                item.LessonsType.Names.Contains(classes.Labs) &&
                                !String.IsNullOrEmpty(classes.Labs) )
                                )
                                {
                                    flag = true;
                                }
                                if ((item.LessonsType.Names == classes.Practics ||
                                item.LessonsType.Names.Contains(classes.Practics) &&
                                !String.IsNullOrEmpty(classes.Practics)))
                                {
                                    flag = true;
                                }
                                
                            }
                        }
                        if (!flag)
                        {
                            classrooms.Add(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }

        }
        public void HighlightLessons(ClassroomsOnLessonsType SelectedClassrooms)
        {
            try
            {
                if (SelectedClassrooms!=null)
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(CheckClassroomsOnLessonsType.TimetableFile))
                    {
                        int listCount = excelPackage.Workbook.Worksheets.Count();
                        List<ExcelWorksheet> anotherWorksheet = new List<ExcelWorksheet>();
                        for (int i = 0; i < listCount; i++)
                        {
                            anotherWorksheet.Add(excelPackage.Workbook.Worksheets[i]);
                        }

                        if (anotherWorksheet[(int)SelectedClassrooms.Page - 1].Cells[(int)SelectedClassrooms.Row, (int)SelectedClassrooms.Col].Value != null)
                        {

                            anotherWorksheet[(int)SelectedClassrooms.Page - 1].Cells[(int)SelectedClassrooms.Row, (int)SelectedClassrooms.Col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                        }
                        excelPackage.SaveAs(CheckLessonsTypeOnEqual.TimetableFile);
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
