using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WpfApp1.Model
{
    public class Window
    {
        public string Day { get; set; }
        public string? Title { get; set; }
        public int LessonNumber { get; set; }
        public GroupsAll Group { get; set; }
        public object Row { get; set; }
        public object Col { get; set; }
        public object Page { get; set; }

        public Window(string day, string? title, int lessonNumber, GroupsAll group,string row,string col,string page)
        {
            Day = day;
            Title = title;
            LessonNumber = lessonNumber;
            Group = group;
            Row = row;
            Col = col;
            Page = page;
        }
        public Window(string day, string? title, int lessonNumber, GroupsAll group, int row, int col, int page)
        {
            Day = day;
            Title = title;
            LessonNumber = lessonNumber;
            Group = group;
            Row = row;
            Col = col;
            Page = page;
        }
    }
    public class CheckTimetableOnWindowViewModel
    {
         List<Window> window { get; set; }
        public ObservableCollection<Window> Windows { get { return windowFromTimetable; } }
        public Window SelectWindows { get; set; }
        private ObservableCollection<Window> windowFromTimetable;
        public RelayCommand HighlightChange { get; set; }
        public CheckTimetableOnWindowViewModel()
        {
            window= new List<Window>();
            windowFromTimetable = new ObservableCollection<Window>();
            HighlightChange = new RelayCommand(o => Highlight(SelectWindows));
            InitIdialWindowsListAsync();
        }
        public async Task InitIdialWindowsListAsync()
        {
            try
            {
                window.Add(new Window("День","Предмет",0,new GroupsAll("Группа",0),"Строка","Столбец","Страница"));
                using (ExcelPackage excelPackage = new ExcelPackage(CheckTimetableOnWindow.TimetableFile))
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
                            for (int j = 8; j < 87; j = j + 2)
                            {

                                if ( item.Cells[j, i].Value != null)
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
                                        int number = 0;
                                        switch (item.Cells[j, 3].Value.ToString())
                                        {
                                            case "1":
                                                number = 1;
                                                break;
                                            case "2":
                                                number = 2;
                                                break;
                                            case "3":
                                                number = 3;
                                                break;
                                            case "4":
                                                number = 4;
                                                break;
                                            case "5":
                                                number = 5;
                                                break;
                                            case "6":
                                                number = 6;
                                                break;
                                            case "7":
                                                number = 7;
                                                break;
                                            default:
                                                break;
                                        }
                                        window.Add(new Window(item.Cells[j-(number-1)*2, 1].Value!=null? item.Cells[j - (number - 1) * 2, 1].Value.ToString():"", item.Cells[j, i].Value.ToString(),number ,new GroupsAll(item.Cells[7, i].Value.ToString(), 0), j, i, item.Index + 1));
                                    }


                                }


                            }

                        }
                    }
                    for (int i = 0; i < window.Count-1; i++)
                    {
                        if (!window[i].Title.Contains("Научно-исследовательская работа") && (!window[i].Title.Contains("День")) ){
                            if (window[i].LessonNumber < window[i + 1].LessonNumber)
                            {
                                if (window[i + 1].LessonNumber - window[i].LessonNumber > 1)
                                {
                                    windowFromTimetable.Add(window[i]);
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
        public void Highlight(Window SelectedClassrooms)
        {
            try
            {
                if (SelectedClassrooms != null)
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(CheckTimetableOnWindow.TimetableFile))
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
                        excelPackage.SaveAs(CheckTimetableOnWindow.TimetableFile);
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
