using System;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Linq;

namespace 시간표
{
    class DataManager
    {
        private readonly string _location;

        public List<string> ListOfDatas = new List<string>();
        public Dictionary<string, List<string>> ListOfSubData = new Dictionary<string, List<string>>();
        public DataManager(string _location)
        {
            this._location = _location;
        }
        public string[] ReadLine(string name, string subAddress)
        {

            var imsi = new List<string>();
            using (var reader = new StreamReader(_location))
            {
                string current;

                while ((current = reader.ReadLine()) != null)
                {
                    if (current.Equals(name + ":"))
                    {
                        while ((current = reader.ReadLine()) != null)
                        {
                            if (current.StartsWith("") && current.EndsWith(":")) break;
                            if (current.Equals("\t-" + subAddress))
                            {
                                while ((current = reader.ReadLine()) != null)
                                {
                                    if (current.StartsWith("") && current.EndsWith(":")) break;
                                    if (current.StartsWith("\t-")) break;
                                    imsi.Add(current.Replace("\t\t-", ""));
                                }

                                break;
                            }
                        }
                        break;
                    }

                }
            }

            var array = imsi.ToArray();
            return array;
        }
        public void ReadAllSubs()
        {
            ListOfSubData.Clear();
            foreach (var name in ListOfDatas)
            {
                using (var reader = new StreamReader(_location))
                {
                    string current;

                    while ((current = reader.ReadLine()) != null)
                    {
                        var imsi = new List<string>();
                        if (!current.Equals(name + ":")) continue;
                        while ((current = reader.ReadLine()) != null)
                        {
                            if (current.StartsWith("") && current.EndsWith(":")) break;
                            if (current.StartsWith("\t-"))
                            {
                                imsi.Add(current.Replace("\t-", ""));
                            }
                        }

                        ListOfSubData.Add(name, imsi);
                        break;

                    }

                }
            }
        }
        public void ReadAllAddress()
        {
            ListOfDatas.Clear();
            using (var reader = new StreamReader(_location))
            {
                string current;
                while ((current = reader.ReadLine()) != null)
                {
                    if (current.StartsWith("") && current.EndsWith(":"))
                    {
                        ListOfDatas.Add(current.Substring(0, current.Length - 1));
                    }
                }
            }
        }

        public void MainAppend(string name, string subAddress, string[] data)
        {
            var datas = new List<string>();
            using (var reader = new StreamReader(_location))
            {
                string current;
                while ((current = reader.ReadLine()) != null)
                {
                    if (current.Equals(name + ":"))
                    {
                        datas.Add(name + ":");
                        datas.Add("\t-" + subAddress);
                        datas.AddRange(data.Select(s => "\t\t-" + s));
                        while ((current = reader.ReadLine()) != null)
                        {
                            if (current.StartsWith("-") || current.StartsWith(""))
                                break;
                        }
                    }

                    datas.Add(current);
                }
            }
            using (var writer = new StreamWriter(_location))
            {
                foreach (var s in datas)
                {
                    writer.WriteLine(s);
                }
            }
        }

        public void SubAppend(string name, string subAddress, string[] data)
        {
            var datas = new List<string>();
            using (StreamReader reader = new StreamReader(_location))
            {
                string current;
                while ((current = reader.ReadLine()) != null)
                {
                    if (current.Equals(name + ":"))
                    {
                        datas.Add(name + ":");
                        datas.Add("\t-" + subAddress);
                        datas.AddRange(data.Select(s => "\t\t-" + s));
                        while ((current = reader.ReadLine()) != null)
                        {
                            if (current.Equals("\t-" + subAddress))
                            {
                                while ((current = reader.ReadLine()) != null)
                                    if (current.StartsWith("\t-") || (current.StartsWith("") && current.EndsWith(":")))
                                        break;
                                break;
                            }
                            datas.Add(current);
                            if (current.StartsWith("") && current.EndsWith(":"))
                                break;
                        }
                    }
                    datas.Add(current);
                }
            }
            using (var writer = new StreamWriter(_location))
            {
                foreach (var s in datas)
                {
                    writer.WriteLine(s);
                }
            }
        }
        public void Append(string name, string subAddress, string[] data)
        {
            if (!IsIn(subAddress, ListOfSubData[name])) MainAppend(name, subAddress, data);
            else SubAppend(name, subAddress, data);
        }

        private bool IsIn(string name, IEnumerable<string> data)
        {
            foreach (var s in ListOfSubData)
            {
                foreach (var str in s.Value) Console.WriteLine("\t" + str);
            }

            return data.Any(s => s == name);
        }
        public void NewAddress(string name, string subAddress, string[] data)
        {
            using (var writer = new StreamWriter(_location, true))
            {
                writer.WriteLine(name + ":");
                writer.WriteLine("\t-" + subAddress);
                foreach (var s in data)
                {
                    writer.WriteLine("\t\t-" + s);
                }
            }
        }
    }
    class FileData
    {
        private readonly string LOC = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\TimeTableData.dat";

        private DataManager dataManager;
        /*
         * initialization
         */
        public FileData()
        {
            if (!File.Exists(LOC)) File.Create(LOC).Close();
            dataManager = new DataManager(LOC);
        }

        public string[] ReadLine(string name, string subAddress)
        {
            return dataManager.ReadLine(name, subAddress);
        }
        public DataManager getManager()
        {
            return this.dataManager;
        }
        public void Write(String name, String subName, String[] data)
        {
            dataManager.ReadAllAddress();
            dataManager.ReadAllSubs();
            if (!isIn(name, dataManager.ListOfDatas))
                dataManager.NewAddress(name, subName, data);
            else dataManager.Append(name, subName, data);
        }

        private bool isIn(string name, List<String> datas)
        {
            foreach (String s in datas)
            {
                if (s.Equals(name)) return true;
            }

            return false;
        }
    }

}
