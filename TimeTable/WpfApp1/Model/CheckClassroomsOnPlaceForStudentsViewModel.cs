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
    public class ClassroomsOnPlace
    {
        public ClassroomsAll Classrooms { get; set; }
        public GroupsAll Group { get; set; }
        public object Row { get; set; }
        public object Col { get; set; }
        public object Page { get; set; }
        public object StudentNumberIdeal { get; set; }
        public object StudentNumberInClass { get; set; }
        public ClassroomsOnPlace(ClassroomsAll classrooms, GroupsAll group, int row, int col, int page, int studentNumberIdeal, int studentNumberInClass)
        {
            Classrooms = classrooms;
            Group = group;
            Row = row;
            Col = col;
            Page = page;
            StudentNumberIdeal = studentNumberIdeal;
            StudentNumberInClass = studentNumberInClass;
        }
        public ClassroomsOnPlace(ClassroomsAll classrooms, GroupsAll group, string row, string col, string page, string studentNumberIdeal, string studentNumberInClass)
        {
            Classrooms = classrooms;
            Group = group;
            Row = row;
            Col = col;
            Page = page;
            StudentNumberIdeal = studentNumberIdeal;
            StudentNumberInClass = studentNumberInClass;
        }
    }
    public class CheckClassroomsOnPlaceForStudentsViewModel
    {
        public ClassroomsFileMeneger classroomsFileMeneger { get; set; }
        public GroupFileMeneger groupsFileMeneger { get; set; }
        private ObservableCollection<ClassroomsOnPlace> classrooms;
        public ClassroomsOnPlace SelectedClassrooms { get; set; }
        private ObservableCollection<ClassroomsAll> classroomsAtFile;
        private ObservableCollection<GroupsAll> gruopsAtFile;
        public ObservableCollection<ClassroomsOnPlace> Classrooms { get { return classrooms; } }
        List<ClassroomsAll> classroomsFromTimetable { get; set; }
        List<GroupsAll> groupsTypeFromTimetable { get; set; }

        public RelayCommand HighlightClassroomsChange { get; set; }
        public List<ClassroomsOnPlace> checkClassrooms { get; set; }
        public CheckClassroomsOnPlaceForStudentsViewModel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            classroomsFileMeneger = new ClassroomsFileMeneger();
            groupsFileMeneger = new GroupFileMeneger();
            classrooms = new ObservableCollection<ClassroomsOnPlace>();
            classroomsAtFile = new ObservableCollection<ClassroomsAll>();
            classroomsFromTimetable = new List<ClassroomsAll>();
            groupsTypeFromTimetable = new List<GroupsAll>();
            gruopsAtFile = new ObservableCollection<GroupsAll>();
            checkClassrooms = new List<ClassroomsOnPlace>();
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
                List<ExcelFile.Group> groupsFromFile = await groupsFileMeneger.Read();
                foreach (var item in groupsFromFile)
                {
                    gruopsAtFile.Add(new GroupsAll(item.Name, item.Cource, item.StudentNumber));
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
                checkClassrooms.Add(new ClassroomsOnPlace(new ClassroomsAll("Аудитория", "", "", 0), new GroupsAll("Группа", 0, 0), "Строка", "Столбец", "Страница","Студентов в группе","Мест в аудитории"));
                using (ExcelPackage excelPackage = new ExcelPackage(CheckClassroomsOnPlaceForStudents.TimetableFile))
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
                            for (int j = 9; j < 88; j = j + 2)
                            {
                                if (item.Cells[j, i].Value != null)
                                {
                                    if (!item.Cells[j , i].Value.ToString().ToLower().Contains("понедельник") &&
                                       !item.Cells[j , i].Value.ToString().ToLower().Contains("вторник") &&
                                       !item.Cells[j, i].Value.ToString().ToLower().Contains("среда") &&
                                       !item.Cells[j, i].Value.ToString().ToLower().Contains("четверг") &&
                                       !item.Cells[j, i].Value.ToString().ToLower().Contains("пятница") &&
                                       !item.Cells[j, i].Value.ToString().ToLower().Contains("суббота") &&
                                       !Regex.IsMatch(item.Cells[j, i].Value.ToString(), @"^[0-9]{3,5}.[0-9]{3,5}", RegexOptions.IgnoreCase) &&
                                       item.Cells[j , i].Value.ToString() != "1" &&
                                       item.Cells[j, i].Value.ToString() != "2" &&
                                       item.Cells[j, i].Value.ToString() != "3" &&
                                       item.Cells[j, i].Value.ToString() != "4" &&
                                       item.Cells[j, i].Value.ToString() != "5" &&
                                       item.Cells[j, i].Value.ToString() != "6" &&
                                       item.Cells[j, i].Value.ToString() != "7")
                                    {
                                        Regex regex = new Regex(@"[1-6].[0-9]{1,3}");
                                        Regex regexSport = new Regex(@"[1-6].[0-9]{1,3}");
                                        MatchCollection matches = regex.Matches(item.Cells[j , i].Value.ToString());
                                        MatchCollection matchesSport = regexSport.Matches(item.Cells[j, i].Value.ToString());
                                        string str = "";
                                        if (matches.Count > 0)
                                        {
                                            int index = item.Cells[j, i].Value.ToString().IndexOf(matches[0].Value[0]);
                                            str = item.Cells[j, i].Value.ToString().Substring(index).Trim();
                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        else if (item.Cells[j, i].Value.ToString().ToLower().Contains("дист"))
                                        {
                                            int index = item.Cells[j, i].Value.ToString().IndexOf("дист");
                                            str = item.Cells[j, i].Value.ToString().Substring(index).Trim();
                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        else if (item.Cells[j, i].Value.ToString().ToLower().Contains("зал"))
                                        {
                                            int index = item.Cells[j , i].Value.ToString().IndexOf("зал");
                                            str = item.Cells[j, i].Value.ToString().Substring(index).Trim();
                                            if (item.Cells[j , i].Value.ToString().Substring(0, index).Trim().Contains("1-") || item.Cells[j, i].Value.ToString().Substring(0, index).Trim().Contains("3-"))
                                            {
                                                str = item.Cells[j, i].Value.ToString().Substring(index - 4).Trim();
                                            }
                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        else if (item.Cells[j , i].Value.ToString().ToLower().Contains("базовая кафедра"))
                                        {
                                            int index = item.Cells[j, i].Value.ToString().IndexOf("базовая кафедра");
                                            str = item.Cells[j, i].Value.ToString().Substring(index).Trim();
                                            classroomsFromTimetable.Add(new ClassroomsAll(str, "", "", 0));
                                        }
                                        if (!item.Cells[7, i].Value.ToString().Contains("время") && !item.Cells[7, i].Value.ToString().Contains("пара") && item.Cells[7, i].Value.ToString() != "1" && item.Cells[7, i].Value.ToString() != "2")
                                        {
                                            bool contains = false;
                                            foreach (var group in groupsTypeFromTimetable)
                                            {
                                                if (group.GroupNames == item.Cells[7, i].Value.ToString())
                                                {
                                                    contains = true;
                                                }
                                            }
                                            if (!contains)
                                            {
                                                if (item.Cells[7, i].Value.ToString().Contains("51") || item.Cells[7, i].Value.ToString().Contains("52"))
                                                {
                                                    groupsTypeFromTimetable.Add(new GroupsAll(item.Cells[7, i].Value.ToString(), 5));
                                                    checkClassrooms.Add(new ClassroomsOnPlace(new ClassroomsAll(String.IsNullOrEmpty(str) ? "нет аудитории" : str, "", "", 0), new GroupsAll(item.Cells[7, i].Value.ToString(), 5), j, i, item.Index + 1,0,0));

                                                }
                                                if (item.Cells[7, i].Value.ToString().Contains("41") || item.Cells[7, i].Value.ToString().Contains("42"))
                                                {
                                                    groupsTypeFromTimetable.Add(new GroupsAll(item.Cells[7, i].Value.ToString(), 4));
                                                    checkClassrooms.Add(new ClassroomsOnPlace(new ClassroomsAll(String.IsNullOrEmpty(str) ? "нет аудитории" : str, "", "", 0), new GroupsAll(item.Cells[7, i].Value.ToString(), 4), j, i, item.Index + 1,0,0));

                                                }
                                                if (item.Cells[7, i].Value.ToString().Contains("31") || item.Cells[7, i].Value.ToString().Contains("32"))
                                                {
                                                    groupsTypeFromTimetable.Add(new GroupsAll(item.Cells[7, i].Value.ToString(), 3));
                                                    checkClassrooms.Add(new ClassroomsOnPlace(new ClassroomsAll(String.IsNullOrEmpty(str) ? "нет аудитории" : str, "", "", 0), new GroupsAll(item.Cells[7, i].Value.ToString(), 3), j, i, item.Index + 1,0,0));

                                                }
                                                if (item.Cells[7, i].Value.ToString().Contains("21") || item.Cells[7, i].Value.ToString().Contains("22"))
                                                {
                                                    groupsTypeFromTimetable.Add(new GroupsAll(item.Cells[7, i].Value.ToString(), 2));
                                                    checkClassrooms.Add(new ClassroomsOnPlace(new ClassroomsAll(String.IsNullOrEmpty(str) ? "нет аудитории" : str, "", "", 0), new GroupsAll(item.Cells[7, i].Value.ToString(), 2), j, i, item.Index + 1,0,0));

                                                }
                                                if (item.Cells[7, i].Value.ToString().Contains("11") || item.Cells[7, i].Value.ToString().Contains("12"))
                                                {
                                                    groupsTypeFromTimetable.Add(new GroupsAll(item.Cells[7, i].Value.ToString(), 1));
                                                    checkClassrooms.Add(new ClassroomsOnPlace(new ClassroomsAll(String.IsNullOrEmpty(str) ? "нет аудитории" : str, "", "", 0), new GroupsAll(item.Cells[7, i].Value.ToString(), 1), j, i, item.Index + 1,0,0));

                                                }
                                            }

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
                            if (item.Classrooms.Names == classes.Names)

                            {
                                item.StudentNumberInClass = classes.PeopleNumber;
                                foreach (var group in gruopsAtFile)
                                {
                                    if (item.Group.GroupNames == group.GroupNames)
                                    {
                                        if ((classes.PeopleNumber>group.StudentNumber)&& group.StudentNumber>0)
                                        {
                                            flag = true;
                                        }
                                        item.StudentNumberIdeal = group.StudentNumber;
                                        
                                    }
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
        public void HighlightLessons(ClassroomsOnPlace SelectedClassrooms)
        {
            try
            {
                if (SelectedClassrooms != null)
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(CheckClassroomsOnPlaceForStudents.TimetableFile))
                    {
                        int listCount = excelPackage.Workbook.Worksheets.Count();
                        List<ExcelWorksheet> anotherWorksheet = new List<ExcelWorksheet>();
                        for (int i = 0; i < listCount; i++)
                        {
                            anotherWorksheet.Add(excelPackage.Workbook.Worksheets[i]);
                        }

                        if (anotherWorksheet[(int)SelectedClassrooms.Page - 1].Cells[(int)SelectedClassrooms.Row, (int)SelectedClassrooms.Col].Value != null)
                        {

                            anotherWorksheet[(int)SelectedClassrooms.Page - 1].Cells[(int)SelectedClassrooms.Row, (int)SelectedClassrooms.Col].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFEA4C89"));
                        }
                        excelPackage.SaveAs(CheckClassroomsOnPlaceForStudents.TimetableFile);
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
