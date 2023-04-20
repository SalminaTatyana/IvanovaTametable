using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp1.Model.ExcelFile;

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
        public  string GroupNames { get; set; }
        public  int Cource { get; set; }
        public  int StudentNumber { get; set; }
        public RelayCommand RemoveGroup { get; set; }

        public GroupsAll(string name, int cource,int count)
        {
            GroupNames = name;
            Cource = cource;
            StudentNumber = count;
        }

    }
    public class MainWinwowViewModal
    {
        private  List<FilesAll> files;
        public  List<FilesAll> Files { get { return files; } }
        private  List<GroupsAll> groups;
        public  List<GroupsAll> Groups { get { return groups; } }
        public  FileManager fileManager = new FileManager();
        public  GroupFileMeneger groupFileManager = new GroupFileMeneger();
        public RelayCommand AddGroup { get; set; }
        public RelayCommand SaveGroupChange { get; set; }
        public RelayCommand DeleteGroupChange { get; set; }


        public MainWinwowViewModal()
        {
            files = new List<FilesAll>();
            groups = new List<GroupsAll>();
            InitOldFilesAsync();
            InitIdialGroupListAsync();
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
                List<Group> file = await groupFileManager.Read();
                foreach (var item in file)
                {
                    groups.Add(new GroupsAll(item.Name, item.Cource,item.StudentNumber));
                }
            }
            catch (Exception ex)
            {

            }

        }
    }
}
