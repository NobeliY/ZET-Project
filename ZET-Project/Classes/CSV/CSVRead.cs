using System;
using System.Collections.Generic;
using System.IO;
using LumenWorks.Framework.IO.Csv;


//using Raschet.Classes.DataBaseScripts;

namespace ZET_Project.CSV
{
    public class CsvRead
    {
        public static List<Person> Personals = new();
        internal static string[] Headers { get; set; }
        private static int Cid { get; set; }
        public static string Post { get; private set; }
        private static bool Filled;
        

        public static void CsvParser(string path, string login, string password)
        {
            Filled = false;
            if (!Filled)
            {
                using var streamReader = new StreamReader(path);
                using CsvReader csvReader = new (streamReader,true);

                Headers = csvReader.GetFieldHeaders();

                while (csvReader.ReadNextRecord())
                {
                    int id = Convert.ToInt32(csvReader[0]);
                    Person item = new()
                    {
                        Id = id, Name = csvReader["NAME"], Surname = csvReader["SURNAME"], Post = csvReader["POST"]
                    };
                    Personals.Add(item);
                
                    // Login and Password Massive
                    var tempLogin = csvReader[4];
                    var tempPassword = csvReader[5];
                
                    if (login.Equals(tempLogin) && password.Equals(tempPassword))
                    {
                        Cid = Convert.ToInt32(csvReader["ID"]);
                        Post = csvReader["POST"];
                        AuthorizationForm.Message =
                            $@"Добро пожаловать, {csvReader["POST"]} {csvReader["NAME"]} {csvReader["SURNAME"]}";

                    }
                }

                Filled = true;
            }
            
        }
        
    }
}