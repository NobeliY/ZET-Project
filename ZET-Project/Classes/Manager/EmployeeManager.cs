using System;
using System.IO;
using System.Threading;
using ZET_Project.Classes.CSV;
using ZET_Project.Classes.Employees;

namespace ZET_Project.Classes.Manager
{
    public static class EmployeeManager
    {
        private const string? Message = "Добро пожаловать, ";
        public static string? Initials { get; set; }

        public static void Start(string? login, string? password)
        {
            
            CsvRead.CsvParser(login,password);
            ExcelManager.SaveExcelFiles();
            switch (CsvRead.Post?.ToLower())
            {
                case "freelancer":
                    Console.Clear();
                    Console.WriteLine($"{Message}, {CsvRead.Post} {Initials}!");
                    Thread.Sleep(1000);
                    Console.ReadLine();
                    Freelancer freelancer = new();
                    freelancer.SendInformation();
                    ExcelManager.SaveExcelFiles();
                    CsvRead.Post = String.Empty;
                    break;
                case "accountant":
                    Console.Clear();
                    Console.WriteLine($"{Message}, {CsvRead.Post} {Initials}!");
                    Thread.Sleep(1000);
                    Accountant accountant = new();
                    accountant.SendInformation();
                    ExcelManager.SaveExcelFiles();
                    CsvRead.Post = String.Empty;
                    break;
                case "director":
                    Console.Clear();
                    Console.WriteLine($"{Message}, {CsvRead.Post} {Initials}!");
                    Thread.Sleep(1000);
                    Director director = new();
                    director.SendInformation();
                    ExcelManager.SaveExcelFiles();
                    CsvRead.Post = String.Empty;
                    break;
            }
        }
    }
}