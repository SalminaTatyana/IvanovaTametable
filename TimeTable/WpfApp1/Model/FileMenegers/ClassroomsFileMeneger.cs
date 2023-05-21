using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp1.Model.ExcelFile;

namespace WpfApp1.Model.FileMenegers
{
    public class ClassroomsFileMeneger
    {
        public string path = "file/classroomFile" + ".txt";
        public async Task Save(List<Classrooms> file)
        {
            if (!File.Exists(path))
            {
                File.CreateText(path);
            }
            // сохранение данных

            using (StreamWriter writer = new StreamWriter(path, false))
            {
                foreach (Classrooms item in file)
                {
                    await writer.WriteLineAsync(item.Names + "|" + item.PeopleNumber.ToString() + "|" + (item.Practics ? 1 : 0).ToString()+"|"+ (item.Labs ? 1 : 0).ToString());
                }
            }
        }
        public async Task<List<Classrooms>> Read()
        {
            List<Classrooms> files = new List<Classrooms>();
            if (!File.Exists(path))
            {
                File.CreateText(path);
            }
            using (StreamReader reader = new StreamReader(path))
            {
                string? line;
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    try {
                        int index = line.IndexOf("|");
                        string name = line.Substring(0, index);
                        int secondIndex = line.Substring(index + 1).IndexOf("|");
                        int thirthIndex = line.Substring(secondIndex + index +2).IndexOf("|");
                        int number = Int32.Parse(line.Substring(index + 1, secondIndex ));
                        string pr = line.Substring(index + secondIndex + 2, thirthIndex);
                        string lb = line.Substring(secondIndex + index + thirthIndex + 3);
                        bool practis =pr.Contains("1"); 
                        bool lab = lb.Contains("1") ; 

                        files.Add(new Classrooms(name,  practis,lab, number));
                    }
                    catch(Exception ex)
                    {

                    }
                    
                }
            }

            return files;
        }
        public void Clear()
        {

            if (!File.Exists(path))
            {
                File.CreateText(path);
            }
            StreamWriter fs = new StreamWriter(path, false);
            fs.Close();
        }
    }
}
