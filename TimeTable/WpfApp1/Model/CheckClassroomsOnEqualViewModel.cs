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
        private ObservableCollection<ClassroomsAll> badClassrooms;
        public ObservableCollection<ClassroomsAll> Classrooms { get { return classrooms; } }
        public ObservableCollection<ClassroomsAll> BadClassrooms { get { return badClassrooms; } }
        public ClassroomsAll SelectedClassrooms { get; set; }
        public ClassroomsAll SelectedBadClassrooms { get; set; }
        public RelayCommand AddClassrooms { get; set; }
        public RelayCommand SaveClassroomsChange { get; set; }
        public RelayCommand HighlightClassroomsChange { get; set; }
        public RelayCommand ReplaceClassroomsChange { get; set; }
        List<ClassroomsAll> lessonsFromTimetable { get; set; }
        public CheckClassroomsOnEqualViewModel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            classroomsFileMeneger = new ClassroomsFileMeneger();
            classrooms = new ObservableCollection<ClassroomsAll>();
            badClassrooms = new ObservableCollection<ClassroomsAll>();
            lessonsFromTimetable = new List<ClassroomsAll>();
            AddClassrooms = new RelayCommand(o => AddNewClassrooms(SelectedBadClassrooms));
            SaveClassroomsChange = new RelayCommand(o => SaveClassrooms());
            HighlightClassroomsChange = new RelayCommand(o => HighlightClassrooms(SelectedBadClassrooms));
            ReplaceClassroomsChange = new RelayCommand(o => ReplaceClassrooms(SelectedClassrooms, SelectedBadClassrooms));
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
        public async Task AddNewClassrooms(ClassroomsAll classes)
        {
           
            bool flag = false;
            for (int i = 0; i < classrooms.Count; i++)
            {
                if (classrooms[i].Names.ToLower() == classes.Names.ToLower())
                {
                    classrooms[i].PeopleNumber = classes.PeopleNumber;
                    flag = true;
                    break;
                }
            }
            if (!flag)
            {
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    classrooms.Add(new ClassroomsAll(classes.Names, classes.Practics,classes.Labs, classes.PeopleNumber));

                });
            }
            await SaveClassrooms();
        }
        public async Task HighlightClassrooms(ClassroomsAll group)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(CheckClassroomOnEqual.TimetableFile))
            {
                int listCount = excelPackage.Workbook.Worksheets.Count();
                List<ExcelWorksheet> anotherWorksheet = new List<ExcelWorksheet>();
                for (int i = 0; i < listCount; i++)
                {
                    anotherWorksheet.Add(excelPackage.Workbook.Worksheets[i]);
                }
                foreach (var item in anotherWorksheet)
                {
                    for (int j = 9; j < 88; j = j + 2)
                    {
                        int col = item.Dimension.End.Column;
                        for (int i = 1; i < col; i++)
                        {

                            if (item.Cells[j, i].Value != null)
                            {
                                if (item.Cells[j, i].Value.ToString().Contains(group.Names))
                                {
                                    item.Cells[j, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGreen);
                                }
                            }
                        }
                    }
                }
                excelPackage.SaveAs(CheckClassroomOnEqual.TimetableFile);
                excelPackage.Dispose();
            }
        }
        public async Task ReplaceClassrooms(ClassroomsAll group, ClassroomsAll badGroup)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(CheckClassroomOnEqual.TimetableFile))
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
                    for (int j = 9; j <88; j=j+2)
                    {
                        for (int i = 1; i < col; i++)
                        {

                            if (item.Cells[j, i].Value != null)
                            {
                                if (item.Cells[j, i].Value.ToString().ToLower().Contains(badGroup.Names.ToLower()))
                                {
                                    item.Cells[j, i].Value = group.Names;
                                }
                            }
                        }
                    }
                    
                }
                excelPackage.SaveAs(CheckClassroomOnEqual.TimetableFile);
                excelPackage.Dispose();
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    badClassrooms.Remove(badGroup);
                });

            }
        }
        public async Task SaveClassrooms()
        {
            List<ExcelFile.Classrooms> saveGroup = new List<ExcelFile.Classrooms>();
            foreach (var group in classrooms)
            {
                saveGroup.Add(new ExcelFile.Classrooms(group.Names, group.Practics.Length>0, group.Labs.Length > 0, group.PeopleNumber));
            }
            App.Current.Dispatcher.Invoke((Action)delegate ()
            {
                classrooms.Clear();
            });
            await classroomsFileMeneger.Save(saveGroup);
            List<Classrooms> file = await classroomsFileMeneger.Read();
            foreach (var item in file)
            {
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    classrooms.Add(new ClassroomsAll(item.Names, "","" ,item.PeopleNumber));
                });
            }
            App.Current.Dispatcher.Invoke((Action)delegate ()
            {
                badClassrooms.Clear();
            });

            foreach (var item in lessonsFromTimetable)
            {
                bool flag = false;
                bool flagBad = false;
                foreach (var group in classrooms)
                {
                    if (item.Names.ToLower() == group.Names.ToLower())
                    {
                        flag = true;
                        break;
                    }
                }
                foreach (var group in badClassrooms)
                {
                    if (item.Names.ToLower() == group.Names.ToLower())
                    {
                        flagBad = true;
                        break;
                    }
                }
                if (!flag&&!flagBad)
                {
                    App.Current.Dispatcher.Invoke((Action)delegate ()
                    {
                        badClassrooms.Add(item);
                    });

                }
            }
        }
        public void GetClassroomsFromTimetable(ExcelPackage excelPackage)
        {
            try
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
                        foreach (var bad in badClassrooms)
                        {
                            if (item.Names == bad.Names)
                            {
                                flag2 = true;
                                break;
                            }
                        }
                        if (!flag2)
                        {
                            badClassrooms.Add(new ClassroomsAll(item.Names, item.Practics, item.Labs, item.PeopleNumber));
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
