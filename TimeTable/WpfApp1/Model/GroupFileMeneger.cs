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
        
            public string path = "file/groupFile" + ".json";
            public async Task Save(List<Group> file)
            {
                if (!File.Exists(path))
                {
                    File.CreateText(path);
                }
                // сохранение данных
                using (StreamWriter fs = new StreamWriter(path, true))
                {
                    fs.WriteLine(JsonConvert.SerializeObject(file).ToString());
                    fs.Close();
                }
            }
            public async Task<string> Read()
            {
                string files = "";
                if (!File.Exists(path))
                {
                    File.CreateText(path);
                }
               BinaryReader br = new BinaryReader(File.OpenRead(path));

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
