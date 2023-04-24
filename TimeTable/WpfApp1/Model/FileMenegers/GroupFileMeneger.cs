using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp1.Model.ExcelFile;

namespace WpfApp1.Model
{
    public class GroupFileMeneger
    {
        
            public string path = "file/groupFile" + ".txt";
            public async Task Save(List<Group> file)
            {
                if (!File.Exists(path))
                {
                    File.CreateText(path);
                }
            // сохранение данных
            
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                foreach (Group group in file)
                {
                    await writer.WriteLineAsync(group.Name + "|"+ group.StudentNumber.ToString());
                }
            }
        }
            public async Task<List<Group>> Read()
            {
                List<Group> files = new List<Group>();
                if (!File.Exists(path))
                {
                    File.CreateText(path);
                }
            using (StreamReader reader = new StreamReader(path))
            {
                string? line;
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    int index = line.IndexOf("|");
                    string name = line.Substring(0,index);
                    int course;
                    int count=0;
                    if (name.Contains("51") || name.Contains("52"))
                    {
                        course = 5;
                    }
                    else if(name.Contains("41") || name.Contains("42"))
                    {
                        course = 4;
                    }
                    else if(name.Contains("31") || name.Contains("32"))
                    {
                        course = 3;
                    }
                    else if(name.Contains("21") || name.Contains("22"))
                    {
                        course = 2;
                    }
                    else
                    {
                        course = 1;
                    }
                    try
                    {
                        count = Int32.Parse(line.Substring(index+1));
                    }
                    catch (Exception ex)
                    {
                    }
                   files.Add(new Group(name, course, count));
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
