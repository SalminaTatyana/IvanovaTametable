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
    public class CheckLessonsTypeOnEqualViewModel
    {
        public LessonsTypeFileMeneger lessonsFileMeneger = new LessonsTypeFileMeneger();
        private ObservableCollection<LessonsType> lessonsType;
        private ObservableCollection<LessonsType> badLessonsType;
        public ObservableCollection<LessonsType> LessonsType { get { return lessonsType; } }
        public ObservableCollection<LessonsType> BadLessonsType { get { return badLessonsType; } }
        public LessonsType SelectedLessonsType { get; set; }
        public LessonsType SelectedBadLessonsType { get; set; }
        public RelayCommand AddLessonsType { get; set; }
        public RelayCommand SaveLessonsTypeChange { get; set; }
        public RelayCommand HighlightLessonsTypeChange { get; set; }
        public RelayCommand ReplaceLessonsTypeChange { get; set; }
        public List<LessonsType> lessonsTypeFromTimetable { get; set; }
        public CheckLessonsTypeOnEqualViewModel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            lessonsFileMeneger = new LessonsTypeFileMeneger();
            lessonsType = new ObservableCollection<LessonsType>();
            badLessonsType = new ObservableCollection<LessonsType>();
            lessonsTypeFromTimetable = new List<LessonsType>();
            AddLessonsType = new RelayCommand(o => AddNewLessonsType(SelectedBadLessonsType));
            SaveLessonsTypeChange = new RelayCommand(o => SaveLessonsType());
            HighlightLessonsTypeChange = new RelayCommand(o => HighlightLessonsType(SelectedBadLessonsType));
            ReplaceLessonsTypeChange = new RelayCommand(o => ReplaceLessonsType(SelectedLessonsType, SelectedBadLessonsType));
            InitIdialLessonsTypeListAsync();

        }
        public async Task InitIdialLessonsTypeListAsync()
        {
            try
            {
                List<string> file = await lessonsFileMeneger.Read();
                foreach (var item in file)
                {
                    lessonsType.Add(new LessonsType(item));
                }
                InitBadLessonsTypeList();
            }
            catch (Exception ex)
            {

            }

        }
        public void InitBadLessonsTypeList()
        {
            using (ExcelPackage excelPackage = new ExcelPackage(CheckLessonsTypeOnEqual.TimetableFile))
            {
                try
                {
                    GetLessonsTypeFromTimetable(excelPackage);
                }
                catch (Exception ex)
                {

                }

            }
        }
        public async Task AddNewLessonsType(LessonsType lesson)
        {

            bool flag = false;
            for (int i = 0; i < lessonsType.Count; i++)
            {
                if (lessonsType[i].Names.ToLower() == lesson.Names.ToLower())
                {
                    flag = true;
                    break;
                }
            }
            if (!flag)
            {
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    lessonsType.Add(new LessonsType(lesson.Names));

                });
            }
            SaveLessonsType();
        }
        public async Task HighlightLessonsType(LessonsType lesson)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(CheckLessonsTypeOnEqual.TimetableFile))
            {
                int listCount = excelPackage.Workbook.Worksheets.Count();
                List<ExcelWorksheet> anotherWorksheet = new List<ExcelWorksheet>();
                for (int i = 0; i < listCount; i++)
                {
                    anotherWorksheet.Add(excelPackage.Workbook.Worksheets[i]);
                }
                foreach (var item in anotherWorksheet)
                {
                    for (int j = 8; j < 87; j = j + 2)
                    {
                        int col = item.Dimension.End.Column;
                        for (int i = 1; i < col; i++)
                        {

                            if (item.Cells[j, i].Value != null)
                            {
                                if (item.Cells[j, i].Value.ToString().Contains(lesson.Names))
                                {
                                    item.Cells[j, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                                }
                            }
                        }
                    }
                }
                excelPackage.SaveAs(CheckLessonsTypeOnEqual.TimetableFile);
                excelPackage.Dispose();
            }
        }
        public async Task ReplaceLessonsType(LessonsType lesson, LessonsType badLesson)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(CheckLessonsTypeOnEqual.TimetableFile))
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
                    for (int j = 8; j < 87; j = j + 2)
                    {
                        for (int i = 1; i < col; i++)
                        {

                            if (item.Cells[j, i].Value != null)
                            {
                                if (item.Cells[j, i].Value.ToString().ToLower().Contains(badLesson.Names.ToLower()))
                                {
                                    item.Cells[j, i].Value = lesson.Names;
                                }
                            }
                        }
                    }

                }
                excelPackage.SaveAs(CheckLessonsTypeOnEqual.TimetableFile);
                excelPackage.Dispose();
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    badLessonsType.Remove(badLesson);
                });

            }
        }
        public async Task SaveLessonsType()
        {
            List<string> saveGroup = new List<string>();
            foreach (var group in lessonsType)
            {
                saveGroup.Add(group.Names);
            }
            App.Current.Dispatcher.Invoke((Action)delegate ()
            {
                lessonsType.Clear();
            });
            await lessonsFileMeneger.Save(saveGroup);
            List<string> file = await lessonsFileMeneger.Read();
            foreach (var item in file)
            {
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    lessonsType.Add(new LessonsType(item));
                });
            }
            App.Current.Dispatcher.Invoke((Action)delegate ()
            {
                badLessonsType.Clear();
            });

            foreach (var item in lessonsTypeFromTimetable)
            {
                bool flag = false;
                foreach (var group in lessonsType)
                {
                    if (item.Names.ToLower() == group.Names.ToLower())
                    {
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                {
                    App.Current.Dispatcher.Invoke((Action)delegate ()
                    {
                        badLessonsType.Add(item);
                    });

                }
            }
        }
        public void GetLessonsTypeFromTimetable(ExcelPackage excelPackage)
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
                        for (int j = 8; j < 87; j = j + 2)
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
                                    Regex regex = new Regex(@"-л.{0,2}$|-п.{0,2}$");
                                    MatchCollection matches = regex.Matches(item.Cells[j, i].Value.ToString());
                                    if (matches.Count > 0)
                                    {
                                        int index = item.Cells[j, i].Value.ToString().IndexOf(matches[0].Value);
                                        string str = item.Cells[j, i].Value.ToString().Substring(index).Trim();
                                        lessonsTypeFromTimetable.Add(new LessonsType(str));
                                    }
                                }

                            }
                        }

                    }
                }
                foreach (var item in lessonsTypeFromTimetable)
                {
                    bool flag = false;
                    foreach (var lessonType in lessonsType)
                    {
                        if (item.Names == lessonType.Names)
                        {
                            flag = true;
                            break;
                        }
                    }
                    if (!flag)
                    {
                        bool flag2 = false;
                        foreach (var bad in badLessonsType)
                        {
                            if (item.Names == bad.Names)
                            {
                                flag2 = true;
                                break;
                            }
                        }
                        if (!flag2)
                        {
                            badLessonsType.Add(new LessonsType(item.Names));
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
