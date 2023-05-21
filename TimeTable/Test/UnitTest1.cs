using OfficeOpenXml;
using WpfApp1.Model;
using WpfApp1.Model.ExcelFile;
using WpfApp1.Model.FileMenegers;

namespace Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public async Task ClassroomsFileMenegerSaveOnNullAsync()
        {
            //arage
            List<Classrooms> files=null;
            ClassroomsFileMeneger testClass = new ClassroomsFileMeneger();
            //act
            await testClass.Save(files);
            //assert

        }
        [TestMethod]
        public async Task ClassroomsFileMenegerSaveOnListNullAsync()
        {
            //arage
            List<Classrooms> files = new List<Classrooms>();
            ClassroomsFileMeneger testClass = new ClassroomsFileMeneger();
            //act
            await testClass.Save(files);
            //assert
        }
        [TestMethod]
        public async Task ClassroomsFileMenegerSaveOnNameNullAsync()
        {
            //arage
            List<Classrooms> files = new List<Classrooms>();
            files.Add(new Classrooms(null, true, true, 1));
            ClassroomsFileMeneger testClass = new ClassroomsFileMeneger();
            //act
            await testClass.Save(files);
            //assert
        }
        [TestMethod]
        public async Task ClassroomsFileMenegerSaveOnPeopleNullAsync()
        {
            //arage
            List<Classrooms> files = new List<Classrooms>();
            files.Add(new Classrooms("", true, true, 0));
            ClassroomsFileMeneger testClass = new ClassroomsFileMeneger();
            //act
            await testClass.Save(files);
            //assert
        }

        [TestMethod]
        public async Task ClassroomsFileMenegerReadAsync()
        {
            //arage
            List<Classrooms> files = new List<Classrooms>();
            files.Add(new Classrooms(null, true, true, 0));
            ClassroomsFileMeneger testClass = new ClassroomsFileMeneger();
            //act
            List<Classrooms> test =await testClass.Read();
            //assert
            Assert.AreEqual(test, test);
        }
        [TestMethod]
        public void ClassroomsFileMenegerClear()
        {
            //arage
            List<Classrooms> files = null;
            ClassroomsFileMeneger testClass = new ClassroomsFileMeneger();
            //act
            testClass.Clear();
            //assert
        }
        [TestMethod]
        public async Task FileMenegerSaveOnNullAsync()
        {
            //arage
            OldFile files = null;
            FileManager testClass = new FileManager();
            //act
            await testClass.Save(files);
            //assert
        }
        [TestMethod]
        public async Task FileMenegerReadAsync()
        {
            //arage
            FileManager testClass = new FileManager();
            //act
            List<OldFile> test = await testClass.Read();
            //assert
            Assert.AreEqual(test, test);
        }
        [TestMethod]
        public  void FileMenegerClearAsync()
        {
            //arage
            FileManager testClass = new FileManager();
            //act
             testClass.Clear();
            //assert
        }
        [TestMethod]
        public async Task GruopFileMenegerSaveOnNullAsync()
        {
            //arage
            List<Group> files = null;
            GroupFileMeneger testClass = new GroupFileMeneger();
            //act
            await testClass.Save(files);
            //assert
        }
        [TestMethod]
        public async Task GruopFileMenegerSaveOnListNullAsync()
        {
            //arage
            List<Group> files = new List<Group>();
            GroupFileMeneger testClass = new GroupFileMeneger();
            //act
            await testClass.Save(files);
            //assert
        }
        [TestMethod]
        public async Task GruopFileMenegerSaveOnNameNullAsync()
        {
            //arage
            List<Group> files = new List<Group>();
            files.Add(new Group(null));
            GroupFileMeneger testClass = new GroupFileMeneger();
            //act
            await testClass.Save(files);
            //assert
        }
        [TestMethod]
        public async Task GruopFileMenegerSaveOnGroupNullAsync()
        {
            //arage
            List<Group> files = new List<Group>();
            files.Add(new Group(null,0));
            GroupFileMeneger testClass = new GroupFileMeneger();
            //act
            await testClass.Save(files);
            //assert
        }
        [TestMethod]
        public async Task GruopFileMenegerReadAsync()
        {
            //arage
            GroupFileMeneger testClass = new GroupFileMeneger();
            //act
            List<Group> test = await testClass.Read();
            //assert
            Assert.AreEqual(test, test);
        }
        [TestMethod]
        public void GruopFileMenegerClear()
        {
            //arage
            GroupFileMeneger testClass = new GroupFileMeneger();
            //act
            testClass.Clear();
            //assert
        }
        [TestMethod]
        public async Task LessonsFileMenegerSaveOnNullAsync()
        {
            //arage
            List<string> files = null;
            LessonsFileMeneger testClass = new LessonsFileMeneger();
            //act
            await testClass.Save(files);
        }
        [TestMethod]
        public async Task LessonsFileMenegerSaveOnListNullAsync()
        {
            List<string> files = new List<string>();
            LessonsFileMeneger testClass = new LessonsFileMeneger();
            //act
            await testClass.Save(files);
        }

        [TestMethod]
        public async Task LessonsFileMenegerReadAsync()
        {
            //arage
            LessonsFileMeneger testClass = new LessonsFileMeneger();
            //act
            List<string> test = await testClass.Read();
            //assert
            Assert.AreEqual(test, test);
        }
        [TestMethod]
        public void LessonsFileMenegerClear()
        {
            //arage
            LessonsFileMeneger testClass = new LessonsFileMeneger();
            //act
            testClass.Clear();
            //assert
        }
        [TestMethod]
        public async Task LessonsTypeFileMenegerSaveOnNullAsync()
        {
            List<string> files = null;
            LessonsTypeFileMeneger testClass = new LessonsTypeFileMeneger();
            //act
            await testClass.Save(files);
        }
        [TestMethod]
        public async Task LessonsTypeFileMenegerSaveOnListNullAsync()
        {
            List<string> files = new List<string>();
            LessonsTypeFileMeneger testClass = new LessonsTypeFileMeneger();
            //act
            await testClass.Save(files);
        }

        [TestMethod]
        public async Task LessonsTypeFileMenegerReadAsync()
        {
            //arage
            LessonsTypeFileMeneger testClass = new LessonsTypeFileMeneger();
            //act
            List<string> test = await testClass.Read();
            //assert
            Assert.AreEqual(test, test);
        }
        [TestMethod]
        public void LessonsTypeFileMenegerClear()
        {
            //arage
            LessonsTypeFileMeneger testClass = new LessonsTypeFileMeneger();
            //act
            testClass.Clear();
            //assert
        }
        [TestMethod]
        public async Task TeachersFileMenegerSaveOnNullAsync()
        {
            List<string> files = null;
            TeachersFileMeneger testClass = new TeachersFileMeneger();
            //act
            await testClass.Save(files);
        }
        [TestMethod]
        public async Task TeachersFileMenegerSaveOnListNullAsync()
        {
            List<string> files = new List<string>();
            TeachersFileMeneger testClass = new TeachersFileMeneger();
            //act
            await testClass.Save(files);
        }

        [TestMethod]
        public async Task TeachersFileMenegerReadAsync()
        {
            //arage
            TeachersFileMeneger testClass = new TeachersFileMeneger();
            //act
            List<string> test = await testClass.Read();
            //assert
            Assert.AreEqual(test, test);
        }
        [TestMethod]
        public void TeachersFileMenegerClear()
        {
            //arage
            TeachersFileMeneger testClass = new TeachersFileMeneger();
            //act
            testClass.Clear();
            //assert
        }
        [TestMethod]
        public async Task ClassroomsInitAsync()
        {
            //arage
            CheckClasscoomsOnLessonsTypeViewModel testClass = new CheckClasscoomsOnLessonsTypeViewModel();
            //act
            await testClass.InitIdialClassroomsListAsync();
            //assert
        }
        [TestMethod]
        public async Task ClassroomsHighlightLessonsAsync()
        {
            //arage
            CheckClasscoomsOnLessonsTypeViewModel testClass = new CheckClasscoomsOnLessonsTypeViewModel();
            ClassroomsOnLessonsType file = new ClassroomsOnLessonsType(null,null,0,0,0);
            //act
            testClass.HighlightLessons(file);
            //assert
        }
        [TestMethod]
        public async Task ClassroomsOnEqualHighlightLessonsAsync()
        {
            //arage
            CheckClassroomsOnEqualViewModel testClass = new CheckClassroomsOnEqualViewModel();
            ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            //act
            await testClass.HighlightClassrooms(file);
            //assert
        }
        public async Task ClassroomsOnEqualInitLessonsAsync()
        {
            //arage
            CheckClassroomsOnEqualViewModel testClass = new CheckClassroomsOnEqualViewModel();
            ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            //act
            await testClass.InitIdialClassroomsListAsync();
            //assert
        }
        public async Task ClassroomsOnEqualInitBadLessonsAsync()
        {
            //arage
            CheckClassroomsOnEqualViewModel testClass = new CheckClassroomsOnEqualViewModel();
            ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            //act
            testClass.InitBadClassroomsList();
            //assert
        }
        [TestMethod]
        public async Task ClassroomsOnEquaAddAsync()
        {
            //arage
            CheckClassroomsOnEqualViewModel testClass = new CheckClassroomsOnEqualViewModel();
            ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            //act
            await testClass.AddNewClassrooms(file);
            //assert
        }
        [TestMethod]
        public async Task ClassroomsOnEquaRepalceAsync()
        {
            //arage
            CheckClassroomsOnEqualViewModel testClass = new CheckClassroomsOnEqualViewModel();
            ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            //act
            await testClass.ReplaceClassrooms(file,file);
            //assert
        }
        [TestMethod]
        public async Task ClassroomsOnEquaSaveAsync()
        {
            //arage
            CheckClassroomsOnEqualViewModel testClass = new CheckClassroomsOnEqualViewModel();
            ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            //act
            await testClass.SaveClassrooms();
            //assert
        }
        [TestMethod]
        public async Task ClassroomsOnEquaGetAsync()
        {
            //arage
            CheckClassroomsOnEqualViewModel testClass = new CheckClassroomsOnEqualViewModel();
            //ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            ExcelPackage file = new ExcelPackage();
            //act
            testClass.GetClassroomsFromTimetable(file);
            //assert
        }
        [TestMethod]
        public async Task ClassroomsPlaceInitAsync()
        {
            //arage
            CheckClassroomsOnPlaceForStudentsViewModel testClass = new CheckClassroomsOnPlaceForStudentsViewModel();
            //act
            await testClass.IdialClassroomsListAsync();
            //assert
        }
        
        [TestMethod]
        public async Task ClassroomsPlaceHighlightLessonsAsync()
        {
            //arage
            CheckClassroomsOnPlaceForStudentsViewModel testClass = new CheckClassroomsOnPlaceForStudentsViewModel();
            ClassroomsOnPlace file = new ClassroomsOnPlace(null, null, 0, 0, 0,0,0);
            //act
            testClass.HighlightLessons(file);
            //assert
        }
        [TestMethod]
        public async Task GroupOnEqualInitBadLessonsAsync()
        {
            //arage
            CheckGroupOnEqualViewModel testClass = new CheckGroupOnEqualViewModel();
            //act
            testClass.InitBadGroupList();
            //assert
        }
        [TestMethod]
        public async Task GroupOnEqualInitLessonsAsync()
        {
            //arage
            CheckGroupOnEqualViewModel testClass = new CheckGroupOnEqualViewModel();
            //act
            testClass.InitIdialGroupListAsync();
            //assert
        }
        [TestMethod]
        public async Task GroupOnEquaAddAsync()
        {
            //arage
            CheckGroupOnEqualViewModel testClass = new CheckGroupOnEqualViewModel();
            GroupsAll file = new GroupsAll(null, 0, 0);
            //act
            await testClass.AddNewGroup(file);
            //assert
        } 
        [TestMethod]
        public async Task GroupOnEquaHilightAsync()
        {
            //arage
            CheckGroupOnEqualViewModel testClass = new CheckGroupOnEqualViewModel();
            GroupsAll file = new GroupsAll(null, 0, 0);
            //act
            await testClass.HighlightGroup(file);
            //assert
        }
        [TestMethod]
        public async Task GroupsOnEquaRepalceAsync()
        {
            //arage
            CheckGroupOnEqualViewModel testClass = new CheckGroupOnEqualViewModel();
            GroupsAll file = new GroupsAll(null, 0, 0);
            //act
            await testClass.ReplaceGroup(file, file);
            //assert
        }
        [TestMethod]
        public async Task GroupsOnEquaSaveAsync()
        {
            //arage
            CheckGroupOnEqualViewModel testClass = new CheckGroupOnEqualViewModel();
            GroupsAll file = new GroupsAll(null, 0, 0);
            //act
            await testClass.SaveGroupsChange();
            //assert
        }
        [TestMethod]
        public async Task GroupOnEquaGetAsync()
        {
            //arage
            CheckGroupOnEqualViewModel testClass = new CheckGroupOnEqualViewModel();
            //ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            ExcelPackage file = new ExcelPackage();
            //act
            testClass.GetGroupFromTimetable(file);
            //assert
        }
        [TestMethod]
        public async Task LessonsOnEqualInitBadLessonsAsync()
        {
            //arage
            CheckLessonsOnEqualViewModel testClass = new CheckLessonsOnEqualViewModel();
            //act
            testClass.InitBadLessonsTypeList();
            //assert
        }
        [TestMethod]
        public async Task LessonsOnEqualInitLessonsAsync()
        {
            //arage
            CheckLessonsOnEqualViewModel testClass = new CheckLessonsOnEqualViewModel();
            //act
            testClass.InitIdialLessonsTypeListAsync();
            //assert
        }
        [TestMethod]
        public async Task LessonsOnEquaAddAsync()
        {
            //arage
            CheckLessonsOnEqualViewModel testClass = new CheckLessonsOnEqualViewModel();
            LessonsAll file = new LessonsAll(null);
            //act
            await testClass.AddNewLessons(file);
            //assert
        }
        [TestMethod]
        public async Task LessonsOnEquaRepalceAsync()
        {
            //arage
            CheckLessonsOnEqualViewModel testClass = new CheckLessonsOnEqualViewModel();
            LessonsAll file = new LessonsAll(null);
            //act
            await testClass.ReplaceLessons(file, file);
            //assert
        }
        [TestMethod]
        public async Task LessonsOnEquaSaveAsync()
        {
            //arage
            CheckLessonsOnEqualViewModel testClass = new CheckLessonsOnEqualViewModel();
            //act
            await testClass.SaveLessons();
            //assert
        }
        [TestMethod]
        public async Task LessonsOnEquaGetAsync()
        {
            //arage
            CheckLessonsOnEqualViewModel testClass = new CheckLessonsOnEqualViewModel();
            //ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            ExcelPackage file = new ExcelPackage();
            //act
            testClass.GetLessonsTypeFromTimetable(file);
            //assert
        }
        [TestMethod]
        public async Task LessonsOnEquaHilightAsync()
        {
            //arage
            CheckLessonsOnEqualViewModel testClass = new CheckLessonsOnEqualViewModel();
            //ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            LessonsAll file = new LessonsAll(null);
            //act
            testClass.HighlightLessons(file);
            //assert
        }
        [TestMethod]
        public async Task LessonsTypeOnEqualInitBadLessonsAsync()
        {
            //arage
            CheckLessonsTypeOnEqualViewModel testClass = new CheckLessonsTypeOnEqualViewModel();
            //act
            testClass.InitBadLessonsTypeList();
            //assert
        }
        [TestMethod]
        public async Task LessonsTypeOnEqualInitLessonsAsync()
        {
            //arage
            CheckLessonsTypeOnEqualViewModel testClass = new CheckLessonsTypeOnEqualViewModel();
            //act
            testClass.InitIdialLessonsTypeListAsync();
            //assert
        }
        [TestMethod]
        public async Task LessonsTypeOnEquaAddAsync()
        {
            //arage
            CheckLessonsTypeOnEqualViewModel testClass = new CheckLessonsTypeOnEqualViewModel();
            LessonsType file = new LessonsType(null);
            //act
            await testClass.AddNewLessonsType(file);
            //assert
        } 
        [TestMethod]
        public async Task LessonsTypeOnEquaHilightAsync()
        {
            //arage
            CheckLessonsTypeOnEqualViewModel testClass = new CheckLessonsTypeOnEqualViewModel();
            LessonsType file = new LessonsType(null);
            //act
            await testClass.HighlightLessonsType(file);
            //assert
        }
        [TestMethod]
        public async Task LessonsTypeOnEquaRepalceAsync()
        {
            //arage
            CheckLessonsTypeOnEqualViewModel testClass = new CheckLessonsTypeOnEqualViewModel();
            LessonsType file = new LessonsType(null);
            //act
            await testClass.ReplaceLessonsType(file, file);
            //assert
        }
        [TestMethod]
        public async Task LessonsTypeOnEquaSaveAsync()
        {
            //arage
            CheckLessonsTypeOnEqualViewModel testClass = new CheckLessonsTypeOnEqualViewModel();
            //act
            await testClass.SaveLessonsType();
            //assert
        }
        [TestMethod]
        public async Task LessonsTypeOnEquaGetAsync()
        {
            //arage
            CheckLessonsTypeOnEqualViewModel testClass = new CheckLessonsTypeOnEqualViewModel();
            //ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            ExcelPackage file = new ExcelPackage();
            //act
            testClass.GetLessonsTypeFromTimetable(file);
            //assert
        }
        [TestMethod]
        public async Task TeachersOnEqualInitBadLessonsAsync()
        {
            //arage
            CheckTeacherEuqalViewModel testClass = new CheckTeacherEuqalViewModel();
            //act
            testClass.InitBadTeachersList();
            //assert
        }
        [TestMethod]
        public async Task TeachersOnEqualInitLessonsAsync()
        {
            //arage
            CheckTeacherEuqalViewModel testClass = new CheckTeacherEuqalViewModel();
            //act
            testClass.InitIdialTeachersListAsync();
            //assert
        }
        [TestMethod]
        public async Task TeachersOnEquaAddAsync()
        {
            //arage
            CheckTeacherEuqalViewModel testClass = new CheckTeacherEuqalViewModel();
            TeachersAll file = new TeachersAll(null);
            //act
            await testClass.AddNewTeachers(file);
            //assert
        }
        [TestMethod]
        public async Task TeachersOnEquaHilightAsync()
        {
            //arage
            CheckTeacherEuqalViewModel testClass = new CheckTeacherEuqalViewModel();
            TeachersAll file = new TeachersAll(null);
            //act
            await testClass.HighlightTeachers(file);
            //assert
        }
        [TestMethod]
        public async Task TeachersOnEquaRepalceAsync()
        {
            //arage
            CheckTeacherEuqalViewModel testClass = new CheckTeacherEuqalViewModel();
            TeachersAll file = new TeachersAll(null);
            //act
            await testClass.ReplaceTeachers(file, file);
            //assert
        }
        [TestMethod]
        public async Task TeachersOnEquaSaveAsync()
        {
            //arage
            CheckTeacherEuqalViewModel testClass = new CheckTeacherEuqalViewModel();
            //act
            await testClass.SaveTeachers();
            //assert
        }
        [TestMethod]
        public async Task TeachersOnEquaGetAsync()
        {
            //arage
            CheckTeacherEuqalViewModel testClass = new CheckTeacherEuqalViewModel();
            //ClassroomsAll file = new ClassroomsAll(null, null, null, 0);
            ExcelPackage file = new ExcelPackage();
            //act
            testClass.GetTeachersFromTimetable(file);
            //assert
        }
        [TestMethod]
        public async Task OnDoubleLessonsInitAsync()
        {
            //arage
            CheckTimetableOnDoubleLessonsInOneClassroomsViewModel testClass = new CheckTimetableOnDoubleLessonsInOneClassroomsViewModel();
            //act
            await testClass.InitIdialClassroomsListAsync();
            //assert
        }
        [TestMethod]
        public async Task OnDoubleLessonsHighlightLessonsAsync()
        {
            //arage
            CheckTimetableOnDoubleLessonsInOneClassroomsViewModel testClass = new CheckTimetableOnDoubleLessonsInOneClassroomsViewModel();
            ClassroomsOnDoubleLessons file = new ClassroomsOnDoubleLessons(null, null, 0, 0, 0,0);
            //act
            testClass.HighlightLessons(file);
            //assert
        }
        [TestMethod]
        public async Task OnLessonsOfTeacherInitAsync()
        {
            //arage
            CheckTimetableOnLessonsOfTeacherViewModel testClass = new CheckTimetableOnLessonsOfTeacherViewModel();
            //act
            await testClass.InitIdialClassroomsListAsync();
            //assert
        }
        [TestMethod]
        public async Task OnLessonsOfTeacherHighlightLessonsAsync()
        {
            //arage
            CheckTimetableOnLessonsOfTeacherViewModel testClass = new CheckTimetableOnLessonsOfTeacherViewModel();
            TeachersOnDoubleLessons file = new TeachersOnDoubleLessons(null, null, 0, 0, 0, 0);
            //act
            testClass.HighlightLessons(file);
            //assert
        }
        [TestMethod]
        public async Task InitOldFiles()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.InitOldFilesAsync();
            //assert
        }
        [TestMethod]
        public async Task InitIdealGroupList()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.InitIdialGroupListAsync();
            //assert
        }
        [TestMethod]
        public async Task InitIdealClassroomList()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.InitIdialClassroomListAsync();
            //assert
        }
        [TestMethod]
        public async Task InitIdealLessonsTypeList()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.InitIdialLessonsTypeListAsync();
            //assert
        }
        [TestMethod]
        public async Task InitIdealLessonsList()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.InitIdialLessonsListAsync();
            //assert
        }
        [TestMethod]
        public async Task InitIdealTeachersList()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.InitIdialTeachersListAsync();
            //assert
        }
        [TestMethod]
        public async Task AddNewGroupOnNull()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.AddNewGroup(null,0);
            //assert
        }
        [TestMethod]
        public async Task AddNewGroupOnEmpty()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.AddNewGroup("", 0);
            //assert
        }
        [TestMethod]
        public async Task AddNewClassroomOnEmpty()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.AddNewClassroom("", 0,false, false);
            //assert
        }
        [TestMethod]
        public async Task AddNewClassroomOnNull()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.AddNewClassroom(null, 0, false, false);
            //assert
        }
        [TestMethod]
        public async Task AddNewLessonsOnEmpty()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.AddNewLessons("");
            //assert
        }
        [TestMethod]
        public async Task AddNewLessonsOnNull()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.AddNewLessons(null);
            //assert
        }
        [TestMethod]
        public async Task AddNewLessonsTypeOnEmpty()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.AddNewLessonsType("");
            //assert
        }
        [TestMethod]
        public async Task AddNewLessonsTypeOnNull()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.AddNewLessonsType(null);
            //assert
        }
        [TestMethod]
        public async Task AddNewTeacherOnEmpty()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.AddNewTeacher("");
            //assert
        }
        [TestMethod]
        public async Task AddNewTeacherOnNull()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.AddNewTeacher(null);
            //assert
        }
        [TestMethod]
        public async Task RemoveTeacherOnNull()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.RemoveTeacher(null);
            //assert
        }
        [TestMethod]
        public async Task RemoveTeacherOnEmpty()
        {
            TeachersAll file = new TeachersAll(null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.RemoveTeacher(file);
            //assert
        }
        [TestMethod]
        public async Task RemoveClassroomOnNull()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.RemoveClassroom(null);
            //assert
        }
        [TestMethod]
        public async Task RemoveClassroomOnEmpty()
        {
            ClassroomsAll file = new ClassroomsAll(null,null,null,0);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.RemoveClassroom(file);
            //assert
        }
        [TestMethod]
        public async Task RemoveGroupOnNull()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.RemoveGroups(null);
            //assert
        }
        [TestMethod]
        public async Task RemoveGroupOnEmpty()
        {
            GroupsAll file = new GroupsAll(null, 0, 0);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.RemoveGroups(file);
            //assert
        }
        [TestMethod]
        public async Task RemoveLessonsOnNull()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.RemoveLesson(null);
            //assert
        }
        [TestMethod]
        public async Task RemoveLessonsOnEmpty()
        {
            LessonsAll file = new LessonsAll(null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.RemoveLesson(file);
            //assert
        }
        [TestMethod]
        public async Task RemoveLessonsTypeOnNull()
        {

            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.RemoveLessonType(null);
            //assert
        }
        [TestMethod]
        public async Task RemoveLessonsTypeOnEmpty()
        {
            LessonsType file = new LessonsType(null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.RemoveLessonType(file);
            //assert
        }
        [TestMethod]
        public async Task DeleteGroupChange()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.DeleteGroupsChange();
            //assert
        }
        [TestMethod]
        public async Task DeleteClassroomsChange()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.DeleteClassroomsChange();
            //assert
        }
        [TestMethod]
        public async Task DeleteTeachersChange()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.DeleteTeacherChange();
            //assert
        }
        [TestMethod]
        public async Task DeleteLessonsChange()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.DeleteLessonChange();
            //assert
        }
        [TestMethod]
        public async Task DeleteLessonsTypeChange()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.DeleteLessonTypeChange();
            //assert
        }
        [TestMethod]
        public async Task SaveGroupChange()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.SaveGroupsChange();
            //assert
        }
        [TestMethod]
        public async Task SaveClassroomsChange()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.SaveClassroomsChange();
            //assert
        }
        [TestMethod]
        public async Task SaveTeachersChange()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.SaveTeacherChange();
            //assert
        }
        [TestMethod]
        public async Task SaveLessonsChange()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.SaveLessonChange();
            //assert
        }
        [TestMethod]
        public async Task SaveLessonsTypeChange()
        {
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.SaveLessonTypeChange();
            //assert
        }
        [TestMethod]
        public async Task OpenChooseFilesOnEmpty()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.OpenChooseFiles(files);
            //assert
        }
        [TestMethod]
        public async Task OpenChooseFilesOnNull()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.OpenChooseFiles(null);
            //assert
        }
        //[TestMethod]
        //public async Task AddFileOnEmpty()
        //{
        //    FilesAll files = new FilesAll(null, null);
        //    //arage
        //    MainWinwowViewModal testClass = new MainWinwowViewModal();
        //    //act
        //    await testClass.AddFile(files);
        //    //assert
        //}
        //[TestMethod]
        //public async Task AddFileOnNull()
        //{
        //    FilesAll files = new FilesAll(null, null);
        //    //arage
        //    MainWinwowViewModal testClass = new MainWinwowViewModal();
        //    //act
        //    await testClass.AddFile(null);
        //    //assert
        //}
        [TestMethod]
        public async Task OpenFileOnEmpty()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.OpenFile("");
            //assert
        }
        [TestMethod]
        public async Task OpenFileOnNull()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            await testClass.OpenFile(null);
            //assert
        }
        [TestMethod]
        public async Task CleanList()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            testClass.CleanList();
            //assert
        }
        [TestMethod]
        public async Task GoCheckClassroomsOnEqual()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
             testClass.GoCheckClassroomsOnEqual();
            //assert
        }
        [TestMethod]
        public async Task GoCheckGroupOnEqual()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            testClass.GoCheckGroupOnEqual();
            //assert
        }
        [TestMethod]
        public async Task GoCheckLessonsOnEqual()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            testClass.GoCheckLessonsOnEqual();
            //assert
        }
        [TestMethod]
        public async Task GoCheckLessonsTypeOnEqual()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            testClass.GoCheckLessonsTypeOnEqual();
            //assert
        }
        [TestMethod]
        public async Task GoCheckTeachOnEqual()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            testClass.GoCheckTeachOnEqual();
            //assert
        }
        [TestMethod]
        public async Task CheckClassroomsOnLessonsType()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            testClass.CheckClassroomsOnLessonsType();
            //assert
        }
        [TestMethod]
        public async Task CheckClassroomsOnPlace()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            testClass.CheckClassroomsOnPlace();
            //assert
        }
        [TestMethod]
        public async Task CheckTimetableOnDoubleLessonsInOne()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            testClass.CheckTimetableOnDoubleLessonsInOne();
            //assert
        }
        [TestMethod]
        public async Task CheckTimetableOnDoubleLessonsOneTeacher()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            testClass.CheckTimetableOnDoubleLessonsOneTeacher();
            //assert
        }
        [TestMethod]
        public async Task CheckTimetableWindow()
        {
            FilesAll files = new FilesAll(null, null);
            //arage
            MainWinwowViewModal testClass = new MainWinwowViewModal();
            //act
            testClass.CheckTimetableWindow();
            //assert
        }
    }
}