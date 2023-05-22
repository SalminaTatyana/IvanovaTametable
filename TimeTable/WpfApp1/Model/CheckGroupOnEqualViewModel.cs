using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp1.Model.ExcelFile;

namespace WpfApp1.Model
{
    public class CheckGroupOnEqualViewModel
    {
        public GroupFileMeneger groupFileManager = new GroupFileMeneger();
        private List<GroupsAll> groups;
        private List<GroupsAll> badGroups;
        public List<GroupsAll> Groups { get { return groups; } }
        public List<GroupsAll> BadGroups { get { return badGroups; } }
        public GroupsAll SelectedBadGroup { get; set; }
        public GroupsAll SelectedGroup { get; set; }
        public RelayCommand AddGroup { get; set; }
        public RelayCommand SaveGroupChange { get; set; }
        public RelayCommand HighlightGroupChange { get; set; }
        public RelayCommand ReplaceGroupChange { get; set; }
        List<GroupsAll> groupFromTimetable { get; set; }

        public CheckGroupOnEqualViewModel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            groupFileManager = new GroupFileMeneger();
            groups = new List<GroupsAll>();
            badGroups = new List<GroupsAll>();
            AddGroup = new RelayCommand(o => AddNewGroup(SelectedBadGroup));
            SaveGroupChange = new RelayCommand(o => SaveGroupsChange());
            HighlightGroupChange = new RelayCommand(o => HighlightGroup(SelectedBadGroup));
            ReplaceGroupChange = new RelayCommand(o => ReplaceGroup(SelectedGroup,SelectedBadGroup));
            groupFromTimetable = new List<GroupsAll>();
            InitIdialGroupListAsync();
        }
        public async Task InitIdialGroupListAsync()
        {
            try
            {
                List<Group> file = await groupFileManager.Read();
                foreach (var item in file)
                {
                    groups.Add(new GroupsAll(item.Name, item.Cource, item.StudentNumber));
                }
                InitBadGroupList();

            }
            catch (Exception ex)
            {

            }

        }
        public void InitBadGroupList()
        {
            using (ExcelPackage excelPackage = new ExcelPackage(CheckGroupOnEqual.TimetableFile))
            {
                try
                {
                    GetGroupFromTimetable(excelPackage);
                }
                catch (Exception ex)
                {

                }

            }
        }
        public async Task AddNewGroup(GroupsAll group)
        {
            if (group!=null)
            {
                if (!String.IsNullOrEmpty(group.GroupNames))
                {
                    int course;
                    if (group.GroupNames.Contains("51") || group.GroupNames.Contains("52"))
                    {
                        course = 5;
                    }
                    else if (group.GroupNames.Contains("41") || group.GroupNames.Contains("42"))
                    {
                        course = 4;
                    }
                    else if (group.GroupNames.Contains("31") || group.GroupNames.Contains("32"))
                    {
                        course = 3;
                    }
                    else if (group.GroupNames.Contains("21") || group.GroupNames.Contains("22"))
                    {
                        course = 2;
                    }
                    else
                    {
                        course = 1;
                    }
                    bool flag = false;
                    for (int i = 0; i < groups.Count; i++)
                    {
                        if (groups[i].GroupNames.ToLower() == group.GroupNames.ToLower())
                        {
                            groups[i].StudentNumber = group.StudentNumber;
                            flag = true;
                            break;
                        }
                    }
                    if (!flag)
                    {
                        App.Current.Dispatcher.Invoke((Action)delegate ()
                        {
                            groups.Add(new GroupsAll(group.GroupNames, course, group.StudentNumber));

                        });
                    }
                    SaveGroupsChange();
                }
            }
           
        }
        public async Task HighlightGroup(GroupsAll group)
        {
            if (group != null)
            {
                if (!String.IsNullOrEmpty(group.GroupNames))
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(CheckGroupOnEqual.TimetableFile))
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

                                if (item.Cells[7, i].Value != null)
                                {
                                    if (item.Cells[7, i].Value.ToString().Contains(group.GroupNames))
                                    {
                                        item.Cells[7, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FF594CEA"));
                                    }
                                }
                            }
                        }
                        excelPackage.SaveAs(CheckGroupOnEqual.TimetableFile);
                        excelPackage.Dispose();
                    }
                }
            }
                 
        }
        public async Task ReplaceGroup(GroupsAll group,GroupsAll badGroup)
        {
            if (group != null&&badGroup!=null)
            {
                if (!String.IsNullOrEmpty(group.GroupNames)&& !String.IsNullOrEmpty(badGroup.GroupNames))
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(CheckGroupOnEqual.TimetableFile))
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

                                if (item.Cells[7, i].Value != null)
                                {
                                    if (item.Cells[7, i].Value.ToString().ToLower().Contains(badGroup.GroupNames.ToLower()))
                                    {
                                        item.Cells[7, i].Value = group.GroupNames;
                                        item.Cells[7, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                                    }
                                }
                            }
                        }
                        excelPackage.SaveAs(CheckGroupOnEqual.TimetableFile);
                        excelPackage.Dispose();
                        App.Current.Dispatcher.Invoke((Action)delegate ()
                        {
                            badGroups.Remove(badGroup);
                        });

                    }
                }
            }
                    
        }
        public async Task SaveGroupsChange()
        {
            List<ExcelFile.Group> saveGroup = new List<ExcelFile.Group>();
            foreach (var group in groups)
            {
                saveGroup.Add(new ExcelFile.Group(group.GroupNames, group.Cource, group.StudentNumber));
            }
            App.Current.Dispatcher.Invoke((Action)delegate ()
            {
                groups.Clear();
            });
            await groupFileManager.Save(saveGroup);
            List<Group> file = await groupFileManager.Read();
            foreach (var item in file)
            {
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    groups.Add(new GroupsAll(item.Name, item.Cource, item.StudentNumber));
                });
            }
            App.Current.Dispatcher.Invoke((Action)delegate ()
            {
                badGroups.Clear();
            });
           
            foreach (var item in groupFromTimetable)
            {
                bool flag = false;
                bool flagBad = false;
                foreach (var group in groups)
                {
                    if (item.GroupNames.ToLower() == group.GroupNames.ToLower())
                    {
                        flag = true;
                        break;
                    }
                }
                foreach (var group in badGroups)
                {
                    if (item.GroupNames.ToLower() == group.GroupNames.ToLower())
                    {
                        flagBad = true;
                        break;
                    }
                }
                if (!flag&&!flagBad)
                {
                    App.Current.Dispatcher.Invoke((Action)delegate ()
                    {
                        badGroups.Add(item);
                    });
                    
                }
            }
           
        }
        public void GetGroupFromTimetable(ExcelPackage excelPackage)
        {
            try
            {
                int listCount=excelPackage.Workbook.Worksheets.Count();
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

                        if (item.Cells[7, i].Value != null && item.Column(i).Width > 0)
                        {
                            if (!item.Cells[7, i].Value.ToString().Contains("время") && !item.Cells[7, i].Value.ToString().Contains("пара") && item.Cells[7, i].Value.ToString() != "1" && item.Cells[7, i].Value.ToString() != "2")
                            {
                                bool contains = false;
                                foreach (var group in groupFromTimetable)
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
                                        groupFromTimetable.Add(new GroupsAll(item.Cells[7, i].Value.ToString(), 5));
                                    }
                                    if (item.Cells[7, i].Value.ToString().Contains("41") || item.Cells[7, i].Value.ToString().Contains("42"))
                                    {
                                        groupFromTimetable.Add(new GroupsAll(item.Cells[7, i].Value.ToString(), 4));
                                    }
                                    if (item.Cells[7, i].Value.ToString().Contains("31") || item.Cells[7, i].Value.ToString().Contains("32"))
                                    {
                                        groupFromTimetable.Add(new GroupsAll(item.Cells[7, i].Value.ToString(), 3));
                                    }
                                    if (item.Cells[7, i].Value.ToString().Contains("21") || item.Cells[7, i].Value.ToString().Contains("22"))
                                    {
                                        groupFromTimetable.Add(new GroupsAll(item.Cells[7, i].Value.ToString(), 2));
                                    }
                                    if (item.Cells[7, i].Value.ToString().Contains("11") || item.Cells[7, i].Value.ToString().Contains("12"))
                                    {
                                        groupFromTimetable.Add(new GroupsAll(item.Cells[7, i].Value.ToString(), 1));
                                    }
                                }
                            }

                        }
                    }
                }
                foreach (var item in groupFromTimetable)
                {
                    bool flag = false;
                    foreach (var group in groups)
                    {
                        if (item.GroupNames.ToLower()==group.GroupNames.ToLower())
                        {
                            flag = true;
                            break;
                        }
                    }
                    if (!flag)
                    {
                        badGroups.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
