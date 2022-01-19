using System;
using System.Threading;
using OfficeOpenXml;
using ZET_Project.Classes.Manager;

namespace ZET_Project
{
    public static class Program
    {
        public static void Main()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Restart:
            Console.Clear();
            Console.WriteLine("Login");
            string? login = Console.ReadLine();
            Console.WriteLine("Password");
            string? password = Console.ReadLine();
            EmployeeManager.Start(login,password);
            Console.WriteLine("Хотите продолжить (1) или заврешить работу программы (0)?");
            try
            {
                if (int.Parse(Console.ReadLine() ?? string.Empty) == 1) goto Restart;
                else Console.WriteLine();
            }
            catch
            {
                Console.WriteLine("Вы ввели не число!");
                Thread.Sleep(200);
                goto Restart;
            }
        }
        
    }
}