using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using LumenWorks.Framework.IO.Csv;
using ZET_Project.Classes.Employees;
using ZET_Project.Classes.Manager;

namespace ZET_Project.Classes.CSV
{
    public static class CsvRead
    {
        private static Dictionary<int,Person> _personals = new();
        private static Dictionary<int, AuthLog> _authLogs = new();
        private static int Cid { get; set; }
        public static string? Post { get; internal set; }
        private static bool _filled = false;
        private static readonly string Path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Classes\Data\Employee.csv");

        internal static Dictionary<int, Person> GetPersonals()
        {
            return _personals;
        }

        internal static int GetLastId()
        {
            return _personals.Count+1;
        }

        public static void CsvParser( string login, string password)
        {
            RST:
            if (!_filled)
            {
                using var streamReader = new StreamReader(Path);
                using CsvReader csvReader = new (streamReader,true);
                string[] headers = csvReader.GetFieldHeaders();
                while (csvReader.ReadNextRecord())
                {
                    var id = Convert.ToInt32(csvReader[0]);
                    Person item = new(csvReader["NAME"], csvReader["SURNAME"], csvReader["POST"]);
                    _personals.Add(id,item);
                    _authLogs.Add(id,new AuthLog(csvReader["LOGIN"], csvReader["PASSWORD"]));
                    
                }
                _filled = true;
                goto RST;
            }
            else
            {
                foreach (var peAuthLog in _authLogs)
                {
                    // Login and Password Massive
                    var tempLogin = peAuthLog.Value.Login;
                    var tempPassword = peAuthLog.Value.Password;
                    if (login != null && password != null)
                    {
                        if (login.Equals(tempLogin) && password.Equals(tempPassword))
                        {
                            Cid = Convert.ToInt32(peAuthLog.Key);
                                if (peAuthLog.Key.Equals(Cid))
                                {
                                    Post = _personals[Cid].Post;
                                    EmployeeManager.Initials = $"{_personals[Cid].Name} {_personals[Cid].Surname}";
                                }
                                
                            
                        }
                    }
                }
            }

            
            
        }

        internal static string GetPost(string employeeNames)
        {
            string[] arrayName = employeeNames.Split(' ');
            foreach (var person in _personals)
            {
                if ((person.Value.Surname.Equals(arrayName[0]) || person.Value.Surname.Equals(arrayName[1])) &&
                    (person.Value.Name.Equals(arrayName[0]) || person.Value.Name.Equals(arrayName[1])))
                {
                    return person.Value.Post;
                }
            }
            return "Not State";
        }

        internal static void AddEmployeeListFile(Dictionary<int, NewEmployee> newEmployees)
        {
            foreach (var (key, value) in newEmployees)
            {
                _personals.Add(key,new Person(value.Name,value.Surname,value.Post));
                _authLogs.Add(key,new AuthLog(value.Login,value.Password));
            }

            using (FileStream fileStream = new FileStream(Path, FileMode.Append, FileAccess.Write))
            {
                foreach (var (key, value) in newEmployees)
                {
                    string newLine = $"\n{key},{value.Name},{value.Surname},{value.Post},{value.Login},{value.Password}";
                    byte[] buff = Encoding.Default.GetBytes(newLine);
                    fileStream.Write(buff,0,buff.Length);
                }
            }
            
        }
        
    }
}