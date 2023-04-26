using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WpfApp1.Model.ExcelFile;
using WpfApp1.Model.FileMenegers;

namespace WpfApp1.Model
{
    public class FilesAll
    {
        public  string FileNames { get; set; }
        public  string FilePaths { get; set; }
        public FilesAll(string name,string path)
        {
            FileNames = name;
            FilePaths = path;
        }

    }
    public class GroupsAll
    {
        public Guid Id { get; set; }
        public string GroupNames { get; set; }
        public  int Cource { get; set; }
        public  int StudentNumber { get; set; }

        public GroupsAll(string name, int cource,int count)
        {
            Id = Guid.NewGuid();
            GroupNames = name;
            Cource = cource;
            StudentNumber = count;
        }
        public GroupsAll(string name, int cource)
        {
            Id = Guid.NewGuid();
            GroupNames = name;
            Cource = cource;
            StudentNumber = 0;
        }
       
    }
    public class TeachersAll
    {
        public Guid Id { get; set; }
        public string Names { get; set; }
        public List<LessonsAll> LesssondNames { get; set; }
        
        
        public TeachersAll(string name)
        {
            Id = Guid.NewGuid();
            Names = name;
        }
        public TeachersAll(string name, List<LessonsAll> lesson)
        {
            Id = Guid.NewGuid();
            Names = name;
            LesssondNames = lesson;
           
        }

    }
    public class LessonsType
    {
        public Guid Id { get; set; }
        public string Names { get; set; }

        

        public LessonsType(string name)
        {
            Id = Guid.NewGuid();
            Names = name;

        }

    }
    public class LessonsAll
    {
        public Guid Id { get; set; }
        public string Names { get; set; }

        

        public LessonsAll(string name)
        {
            Id = Guid.NewGuid();
            Names = name;

        }
    }
    public class ClassroomsAll
    {
        public Guid Id { get; set; }
        public string Names { get; set; }
        public int PeopleNumber { get; set; }
        public string Practics { get; set; }
        public string Labs { get; set; }

        public ClassroomsAll(string name, string practics, string labs, int peopleNumber)
        {
            Id = Guid.NewGuid();
            Names = name;
            Practics = practics;
            Labs = labs;
            PeopleNumber = peopleNumber;
        }
    }
    public class MainWinwowViewModal
    {
        private ObservableCollection<FilesAll> files;
        public ObservableCollection<FilesAll> Files { get { return files; } }
        private ObservableCollection<GroupsAll> groups;
        public ObservableCollection<GroupsAll> Groups { get { return groups; } }
        public string GroupName { get; set; }
        public int GroupCount { get; set; }
        public Guid GroupId { get; set; }
        private ObservableCollection<LessonsAll> lessons;
        public ObservableCollection<LessonsAll> Lessons { get { return lessons; } }
        public string LessonName { get; set; } 
        public Guid LessonId { get; set; } 
        private ObservableCollection<ClassroomsAll> classes;
        public ObservableCollection<ClassroomsAll> Classes { get { return classes; } }
        public string ClasseName { get; set; }
        public int ClassCount { get; set; }
        public Guid ClassId { get; set; }
        public bool Lab { get; set; }
        public bool Practice { get; set; }

        private ObservableCollection<LessonsType> lessonsType;
        public ObservableCollection<LessonsType> LessonsType { get { return lessonsType; } }
        public string LessonTypeName { get; set; }
        public Guid LessonTypeId { get; set; }

        private ObservableCollection<TeachersAll> teachers;
        public ObservableCollection<TeachersAll> Teachers { get { return teachers; } }
        public string TeacherName { get; set; }
        public Guid TeacherId { get; set; }

        public FileManager fileManager = new FileManager();
        public  GroupFileMeneger groupFileManager = new GroupFileMeneger();
        public  ClassroomsFileMeneger classFileManager = new ClassroomsFileMeneger();
        public  LessonsTypeFileMeneger lessonsTypeFileManager =new LessonsTypeFileMeneger();
        public  LessonsFileMeneger lessonsFileManager =new LessonsFileMeneger();
        public  TeachersFileMeneger teachersFileManager =new TeachersFileMeneger();
        public RelayCommand AddGroup { get; set; }
        public RelayCommand RemoveGroup { get; set; }
        public RelayCommand RemoveLessons { get; set; }
        public RelayCommand RemoveLessonsType { get; set; }
        public RelayCommand RemoveClassrooms { get; set; }
        public RelayCommand SaveGroupChange { get; set; }
        public RelayCommand DeleteGroupChange { get; set; }
        public RelayCommand AddLessons { get; set; }
        public RelayCommand SaveLessonsChange { get; set; }
        public RelayCommand DeleteLessonsChange { get; set; }
        public RelayCommand AddLessonsType { get; set; }
        public RelayCommand SaveLessonsTypeChange { get; set; }
        public RelayCommand DeleteLessonsTypeChange { get; set; }
        public RelayCommand AddTeachers { get; set; }
        public RelayCommand RemoveTeachers { get; set; }
        public RelayCommand SaveTeachersChange { get; set; }
        public RelayCommand DeleteTeachersChange { get; set; }
        public RelayCommand AddClassroom{ get; set; }
        public RelayCommand SaveClassroomChange { get; set; }
        public RelayCommand DeleteClassroomChange { get; set; }
        public RelayCommand OpenChooseFile { get; set; }
        public RelayCommand AddFiles { get; set; }
        public RelayCommand OpenCheckGroup { get; set; }
        public RelayCommand OpenCheckTeach { get; set; }
        public RelayCommand OpenCheckClassroom { get; set; }
        public RelayCommand OpenCheckLessonsType { get; set; }
        public RelayCommand OpenCheckLessons { get; set; }
        public RelayCommand CleanFiles { get; set; }
        public GroupsAll SelectedItem { get; set; }
        public TeachersAll SelectedTeachers { get; set; }
        public LessonsAll SelectedLessons { get; set; }
        public LessonsType SelectedLessonsType { get; set; }
        public ClassroomsAll SelectedClassrooms { get; set; }
        public FilesAll SelectedFile { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;
        public ObservableCollection<string> NowFileNameChoose { get { return nowFileNameChoose; } }
        private ObservableCollection<string> nowFileNameChoose;
        public string SelectedNowFile { get; set; }


        public MainWinwowViewModal()
        {
            files = new ObservableCollection<FilesAll>();
            groups = new ObservableCollection<GroupsAll>();
            classes = new ObservableCollection<ClassroomsAll>();
            teachers = new ObservableCollection<TeachersAll>();
            lessons = new ObservableCollection<LessonsAll>();
            lessonsType = new ObservableCollection<LessonsType>();
            nowFileNameChoose = new ObservableCollection<string>();
            AddGroup = new RelayCommand(o =>AddNewGroup(GroupName,GroupCount));
            AddLessons = new RelayCommand(o => AddNewLessons(LessonName));
            AddLessonsType = new RelayCommand(o => AddNewLessonsType(LessonTypeName));
            AddTeachers = new RelayCommand(o => AddNewTeacher(TeacherName));
            AddClassroom = new RelayCommand(o => AddNewClassroom(ClasseName,ClassCount,Practice,Lab));
            RemoveGroup = new RelayCommand(o => RemoveGroups(SelectedItem));
            RemoveTeachers = new RelayCommand(o => RemoveTeacher(SelectedTeachers));
            RemoveLessons = new RelayCommand(o => RemoveLesson(SelectedLessons));
            RemoveLessonsType = new RelayCommand(o => RemoveLessonType(SelectedLessonsType));
            RemoveClassrooms = new RelayCommand(o => RemoveClassroom(SelectedClassrooms));
            DeleteGroupChange= new RelayCommand(o => DeleteGroupsChange());
            DeleteLessonsTypeChange = new RelayCommand(o => DeleteLessonTypeChange());
            DeleteLessonsChange = new RelayCommand(o => DeleteLessonChange());
            DeleteTeachersChange = new RelayCommand(o => DeleteTeacherChange());
            DeleteClassroomChange = new RelayCommand(o => DeleteClassroomsChange());
            SaveGroupChange= new RelayCommand(o => SaveGroupsChange());
            SaveLessonsChange= new RelayCommand(o => SaveLessonChange());
            SaveLessonsTypeChange= new RelayCommand(o => SaveLessonTypeChange());
            SaveTeachersChange= new RelayCommand(o => SaveTeacherChange());
            SaveClassroomChange= new RelayCommand(o => SaveClassroomsChange());
            OpenChooseFile = new RelayCommand(o => OpenChooseFiles(SelectedFile));
            AddFiles = new RelayCommand(o => AddFile(SelectedFile));
            CleanFiles = new RelayCommand(o => CleanList());
            OpenCheckGroup = new RelayCommand(o => GoCheckGroupOnEqual());
            OpenCheckTeach = new RelayCommand(o => GoCheckTeachOnEqual());
            OpenCheckClassroom = new RelayCommand(o => GoCheckClassroomsOnEqual());
            OpenCheckLessonsType = new RelayCommand(o => GoCheckLessonsTypeOnEqual());
            OpenCheckLessons = new RelayCommand(o => GoCheckLessonsOnEqual());
            InitOldFilesAsync();
            InitIdialGroupListAsync();
            InitIdialClassroomListAsync();
            InitIdialLessonsTypeListAsync();
            InitIdialTeachersListAsync();
            InitIdialLessonsListAsync();
        }
        public  async Task InitOldFilesAsync()
        {
            try
            {
                List < OldFile > file = await fileManager.Read();
                foreach (var item in file)
                {
                    files.Add(new FilesAll(Path.GetFileName(item.Name), item.Name));
                   
                }
            }
            catch (Exception ex)
            {

            }
        }
        public  async Task InitIdialGroupListAsync()
        {
            try
            {
                List<ExcelFile.Group> file = await groupFileManager.Read();
                foreach (var item in file)
                {
                    groups.Add(new GroupsAll(item.Name, item.Cource,item.StudentNumber));
                }
            }
            catch (Exception ex)
            {

            }
        }
        public  async Task InitIdialClassroomListAsync()
        {
            try
            {
                List<Classrooms> file = await classFileManager.Read();
                foreach (var item in file)
                {
                    classes.Add(new ClassroomsAll(item.Names, item.Practics?"пр":"",item.Labs ? "лб" : "", item.PeopleNumber));
                }
            }
            catch (Exception ex)
            {

            }
        }
        public async Task InitIdialLessonsTypeListAsync()
        {
            try
            {
                List<string> file = await lessonsTypeFileManager.Read();
                foreach (var item in file)
                {
                    lessonsType.Add(new LessonsType(item));
                }
            }
            catch (Exception ex)
            {

            }

        }
        public async Task InitIdialLessonsListAsync()
        {
            try
            {
                List<string> file = await lessonsFileManager.Read();
                foreach (var item in file)
                {
                    lessons.Add(new LessonsAll(item));
                }
            }
            catch (Exception ex)
            {

            }

        }
        public async Task InitIdialTeachersListAsync()
        {
            try
            {
                List<string> file = await teachersFileManager.Read();
                foreach (var item in file)
                {
                    teachers.Add(new TeachersAll(item));
                }
            }
            catch (Exception ex)
            {

            }

        }

        public async Task AddNewGroup(string name,int count)
        {
            int course;
            if (name.Contains("51") || name.Contains("52"))
            {
                course = 5;
            }
            else if (name.Contains("41") || name.Contains("42"))
            {
                course = 4;
            }
            else if (name.Contains("31") || name.Contains("32"))
            {
                course = 3;
            }
            else if (name.Contains("21") || name.Contains("22"))
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
                if (groups[i].GroupNames.ToLower()==name.ToLower())
                {
                    groups[i].StudentNumber = count;
                    flag = true;
                    break;
                }
            }
            if (!flag)
            {
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    groups.Add(new GroupsAll(name, course, count));
                });
            } 
        }

        public async Task RemoveGroups(GroupsAll Index)
        {
            groups.Remove(Index);
        } 
        public async Task RemoveTeacher(TeachersAll Index)
        {
            teachers.Remove(Index);
        }
        public async Task RemoveLesson(LessonsAll Index)
        {
            lessons.Remove(Index);
        } 
        public async Task RemoveLessonType(LessonsType Index)
        {
            lessonsType.Remove(Index);
        } 
        public async Task RemoveClassroom(ClassroomsAll Index)
        {
            classes.Remove(Index);
        }
        public async Task DeleteGroupsChange()
        {
            groups.Clear();
            await InitIdialGroupListAsync();
        } 
        public async Task DeleteLessonChange()
        {
            lessons.Clear();
            await InitIdialLessonsListAsync();
        }
        public async Task DeleteLessonTypeChange()
        {
            lessonsType.Clear();
            await InitIdialLessonsTypeListAsync();
        } 
        public async Task DeleteTeacherChange()
        {
            teachers.Clear();
            await InitIdialTeachersListAsync();
        } 
        public async Task DeleteClassroomsChange()
        {
            classes.Clear();
            await InitIdialClassroomListAsync();
        }
        public async Task SaveGroupsChange()
        {
           List<ExcelFile.Group> saveGroup = new List<ExcelFile.Group>();
            foreach (var group in groups)
            {
                saveGroup.Add(new ExcelFile.Group(group.GroupNames,group.Cource,group.Cource));
            }
            await groupFileManager.Save(saveGroup);
        }
        public async Task SaveLessonChange()
        {
            List<string> saveLesson = new List<string>();
            foreach (var lesson in lessons)
            {
                saveLesson.Add(lesson.Names);
            }
            await lessonsFileManager.Save(saveLesson);
        }
        public async Task SaveLessonTypeChange()
        {
            List<string> saveLesson = new List<string>();
            foreach (var lesson in lessonsType)
            {
                saveLesson.Add(lesson.Names);
            }
            await lessonsTypeFileManager.Save(saveLesson);
        }
        public async Task SaveTeacherChange()
        {
            List<string> teacherList = new List<string>();
            foreach (var teacher in teachers)
            {
                teacherList.Add(teacher.Names);
            }
            await teachersFileManager.Save(teacherList);
        }
        public async Task SaveClassroomsChange()
        {
            List<Classrooms> saveClassroom = new List<Classrooms>();
            foreach (var room in classes)
            {
                saveClassroom.Add(new Classrooms(room.Names,!String.IsNullOrEmpty(room.Practics), !String.IsNullOrEmpty(room.Labs),room.PeopleNumber));
            }
            await classFileManager.Save(saveClassroom);
        }
        public async Task AddNewLessons(string name)
        {
           
            bool flag = false;
            for (int i = 0; i < lessons.Count; i++)
            {
                if (lessons[i].Names.ToLower() == name.ToLower())
                {
                    flag = true;
                    break;
                }
            }
            if (!flag)
            {
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    lessons.Add(new LessonsAll(name));
                });
            }
        }
        public async Task OpenChooseFiles(FilesAll file)
        {
            await OpenFile(file.FilePaths);
        }
        public async Task AddFile(FilesAll file)
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
                var read = fileManager.Read().Result;
                for (int i = 0; i < read.Count(); i++)
                {
                    if (read[i].Path == newFile.Path)
                    {
                        contains = true;
                    }
                }
                if (!contains)
                {
                    await fileManager.Save(newFile);
                    App.Current.Dispatcher.Invoke((Action)delegate ()
                    {
                        files.Add(new FilesAll(Path.GetFileName(newFile.Name), newFile.Path));
                    });

                }
                
                await OpenFile(path);
            }
        }

        public async Task AddNewLessonsType(string name)
        {

            bool flag = false;
            for (int i = 0; i < lessonsType.Count; i++)
            {
                if (lessonsType[i].Names.ToLower() == name.ToLower())
                {
                    flag = true;
                    break;
                }
            }
            if (!flag)
            {
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    lessonsType.Add(new LessonsType(name));
                });
            }
        }
       
        public async Task AddNewTeacher(string name)
        {

            bool flag = false;
            for (int i = 0; i < teachers.Count; i++)
            {
                if (teachers[i].Names.ToLower() == name.ToLower())
                {
                    flag = true;
                    break;
                }
            }
            if (!flag)
            {
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    teachers.Add(new TeachersAll(name));
                });
            }
        }
    
        public async Task AddNewClassroom(string name, int count,bool practice, bool lab)
        {

            bool flag = false;
            for (int i = 0; i < groups.Count; i++)
            {
                if (classes[i].Names.ToLower() == name.ToLower())
                {
                    classes[i].PeopleNumber = count;
                    classes[i].Practics = practice?"Пр":"";
                    classes[i].Labs = lab? "Лб" : "";
                    flag = true;
                    break;
                }
            }
            if (!flag)
            {
                App.Current.Dispatcher.Invoke((Action)delegate ()
                {
                    classes.Add(new ClassroomsAll(name, practice ? "Пр" : "", lab ? "Лб" : "", count));
                });
            }
        }
        public async Task OpenFile(string path)
        {
            
            App.Current.Dispatcher.Invoke((Action)delegate ()
            {
                nowFileNameChoose.Add(path);
            });
            MainWindow.timetableFile = new FileInfo(path);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }
        public void CleanList()
        {
            fileManager.Clear();
            App.Current.Dispatcher.Invoke((Action)delegate ()
            {
                files.Clear();
            });
        }
        public void GoCheckGroupOnEqual()
        {
            if (SelectedFile != null)
            {
                CheckGroupOnEqual view = new CheckGroupOnEqual(new FileInfo(SelectedFile.FilePaths));
                view.Show();
            }

        }
        public void GoCheckTeachOnEqual()
        {
            if (SelectedFile != null)
            {
                CheckedTeachersOnEqual view = new CheckedTeachersOnEqual(new FileInfo(SelectedFile.FilePaths));
                view.Show();
            }

        } 
        public void GoCheckClassroomsOnEqual()
        {
            if (SelectedFile != null)
            {
                CheckClassroomOnEqual view = new CheckClassroomOnEqual(new FileInfo(SelectedFile.FilePaths));
                view.Show();
            }

        } 
        public void GoCheckLessonsTypeOnEqual()
        {
            if (SelectedFile != null)
            {
                CheckLessonsTypeOnEqual view = new CheckLessonsTypeOnEqual(new FileInfo(SelectedFile.FilePaths));
                view.Show();
            }

        }
        public void GoCheckLessonsOnEqual()
        {
            if (SelectedFile != null)
            {
                CheckLessonsOnEqual view = new CheckLessonsOnEqual(new FileInfo(SelectedFile.FilePaths));
                view.Show();
            }

        }
    }
}
