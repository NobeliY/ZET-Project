using System;
using System.IO;
using System.Threading;
using OfficeOpenXml;
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
            
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Classes\Data\Employee.csv");
            CsvRead.CsvParser(path,login,password);
            ExcelManager.SaveExcelFiles();
            switch (CsvRead.Post?.ToLower())
            {
                case "freelancer":
                    Console.Clear();
                    Console.WriteLine($"{Message}, {CsvRead.Post} {Initials}!");
                    Thread.Sleep(200);
                    Console.ReadLine();
                    Freelancer freelancer = new();
                    freelancer.SendInformation();
                    ExcelManager.SaveExcelFiles();
                    CsvRead.Post = String.Empty;
                    break;
                case "accountant":
                    Console.Clear();
                    Console.WriteLine($"{Message}, {CsvRead.Post} {Initials}!");
                    Thread.Sleep(200);
                    Accountant accountant = new();
                    accountant.SendInformation();
                    ExcelManager.SaveExcelFiles();
                    CsvRead.Post = String.Empty;
                    break;
                case "director":
                    Console.Clear();
                    Console.WriteLine($"{Message}, {CsvRead.Post} {Initials}!");
                    Thread.Sleep(200);
                    Director director = new();
                    director.SendInformation();
                    ExcelManager.SaveExcelFiles();
                    CsvRead.Post = String.Empty;
                    break;
            }
        }
    }
}