using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using OfficeOpenXml;
namespace ZET_Project.Classes.Manager
{
    public class ExcelPerson
    {
        protected internal int hours;
        protected internal string? note;
        protected internal string? date;

        public ExcelPerson(int hours, string? note, string? date)
        {
            this.hours = hours;
            this.note = note;
            this.date = date;
        }

        
    }

    public class ExcelPersonAll
    {
        protected internal int hours;
        protected internal string? note;
        protected internal string? date;
        protected internal string employeeName;
        public ExcelPersonAll(string employeeName, int hours, string? note, string? date)
        {
            this.hours = hours;
            this.date = date;
            this.note = note;
            this.employeeName = employeeName;
        }
        
    }
    public class ExcelManager
    {
#pragma warning disable 8714
        protected internal static Dictionary<string, ExcelPerson> ExcelPersons =
            new();
        protected internal static Dictionary<int, ExcelPersonAll> ExcelPersonAlls = 
            new();
#pragma warning restore 8714
        public static string path = @"..\..\..\Classes\Data\EmployeeReports.xlsx";

        public static void SaveExcelFiles()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            for (int i = 0; i <= 2; i++)
            {
                using var src = new ExcelPackage(new FileInfo(path));
                using var dest = new ExcelPackage(new FileInfo($"Employee{src.Workbook.Worksheets[i].Name}.xlsx"));
                var wsSrc = src.Workbook.Worksheets[i];
                var wsDest = dest.Workbook.Worksheets[wsSrc.Name] ?? dest.Workbook.Worksheets.Add(wsSrc.Name);
                for (var r = 1; r <= wsSrc.Dimension.Rows; r++)
                {
                    for (var c = 1; c <= wsSrc.Dimension.Columns; c++) 
                    {
                        var cellSrc = wsSrc.Cells[r, c];
                        var cellDest = wsDest.Cells[r, c];
                        // Copy value
                        cellDest.Value = cellSrc.Value;
                    }
                }
                dest.Save();
            }
        }

        public void GetReportEmployeeArray(string employeeName, string tableList)
        {
            var package = new ExcelPackage(path);
            int lastRow = 0;
            if ((employeeName.Contains("All") || employeeName.Contains("Все")) && tableList.Contains("All"))
            {
                int Cid = 0;
                foreach (var sWorksheet in package.Workbook.Worksheets)
                {
                    lastRow = sWorksheet.Dimension.End.Row;
                    while (sWorksheet.Cells[lastRow,1].Value == null)
                    {
                        lastRow--;
                    }
                    
                    for (int i = 2; i <= lastRow; i++)
                    {
                        try
                        {
                            ExcelPersonAlls.Add(Cid, new ExcelPersonAll(sWorksheet.Cells[$"B{i}"].Text, 
                                sWorksheet.Cells[$"C{i}"].GetValue<int>(),
                                note: sWorksheet.Cells[$"D{i}"].Text,
                                date: sWorksheet.Cells[$"A{i}"].Text));
                        }
                        catch
                        {
                            /*ExcelPersons[sWorksheet.Cells[$"B{i}"].Text].hours +=
                                sWorksheet.Cells[$"C{i}"].GetValue<int>();*/
                            Console.WriteLine("Done");
                        }
                        Cid++;
                    }  
                }
                
            }
            else
            {
                var sheet = package.Workbook.Worksheets[tableList];
                lastRow = sheet.Dimension.End.Row;
                while (sheet.Cells[lastRow,1].Value == null)
                {
                    lastRow--;
                }

                for (int i = 2; i <= lastRow; i++)
                {
                    if (sheet.Cells[i,2].Value.Equals(employeeName))
                    {
                        ExcelPersons.Add(sheet.Cells[$"A{i}"].Text, new ExcelPerson(sheet.Cells[$"C{i}"].GetValue<int>(),
                                note: sheet.Cells[$"D{i}"].Text,
                                date: sheet.Cells[$"B{i}"].Text));
                    }
                }    
            }
            
        }
        public void AddHours(string? initials, string? date, int hours, string? note, string? tableList)
        {
            var package = new ExcelPackage(path);
            var sheet = package.Workbook.Worksheets[tableList];
            int lastrow = sheet.Dimension.End.Row;
            while (sheet.Cells[lastrow,1].Value == null)
            {
                lastrow--;
            }

            for (int i = 2; i <= lastrow; i++)
            {
                if (sheet.Cells[i,1].Text.Equals(date))
                {
                    sheet.Cells[i, 3].Value = sheet.Cells[i,3].GetValue<Int32>() + hours;
                }
                else
                {
                    if (i == lastrow)
                    {
                        sheet.Cells[i + 1, 1].Value = date; // Дата (Например. 11.01.2021)
                        sheet.Cells[i + 1, 2].Value = initials;     // Фамилия и Имя сотрудника
                        sheet.Cells[i + 1, 3].Value = hours;        // Количество добавленных часов
                        sheet.Cells[i + 1, 4].Value = note;         // За что добавили часы
                    }
                }
            }

            package.Save();
        }
        
    }
    public class Report
    {
        public DateTime DateTime { get; set; }
        public string? Initials { get; set; }
        public int Hours { get; set; }
        public string? Note { get; set; }
    }
    public class ReportExcelGenerator
    {
        private static string? _worksheet;

        private Report GetReport()
        {
            return new Report()
            {
                DateTime = new DateTime(2022, 1, 1), Initials = EmployeeManager.Initials, Hours = 0, Note = "Test"
            };
        }

        private byte[] Generate(Report report)
        {
            var package = new ExcelPackage();
            if (!File.Exists(ExcelManager.path))
            {
                var sheet = package.Workbook.Worksheets.Add("Образец");
                sheet.Cells["A1"].Value = "Дата";
                sheet.Cells["B1"].Value = "Фамилия и Имя сотрудника";
                sheet.Cells["C1"].Value = "Часы";
                sheet.Cells["D1"].Value = "Примечание";
                var sheetD = package.Workbook.Worksheets.Add("Director", sheet);
                var sheetF = package.Workbook.Worksheets.Add("Freelancer", sheet);
                var sheetA = package.Workbook.Worksheets.Add("Accountant", sheet);
                package.Save();
            }

            if (package.Workbook.Worksheets[_worksheet] != null)
            {
                var sheet = package.Workbook.Worksheets.Add(_worksheet);
                sheet.Cells["A1"].Value = "Дата";
                sheet.Cells["B1"].Value = "Фамилия и Имя сотрудника";
                sheet.Cells["C1"].Value = "Часы";
                sheet.Cells["D1"].Value = "Примечание";
            }
            return package.GetAsByteArray();
        }

        public static void Create(string sheet)
        {
            _worksheet = sheet;
            var reportData = new ReportExcelGenerator().GetReport();
            var reportExcel = new ReportExcelGenerator().Generate(reportData);
            File.WriteAllBytes(ExcelManager.path, reportExcel);
        }
    }
}