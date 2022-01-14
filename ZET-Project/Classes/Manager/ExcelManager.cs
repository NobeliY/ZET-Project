using System;
using System.IO;
using System.Runtime.CompilerServices;
using OfficeOpenXml;
namespace ZET_Project.Classes.Manager
{
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
            if (!File.Exists(@"..\..\..\Classes\Data\EmployeeReports.xlsx"))
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "Date";
                sheet.Cells["B1"].Value = "Initials";
                sheet.Cells["C1"].Value = "Hours";
                sheet.Cells["D1"].Value = "Note";
            }

            if (!package.Workbook.Worksheets.Equals(_worksheet))
            {
                var sheet = package.Workbook.Worksheets.Add(_worksheet);
                sheet.Cells["A1"].Value = "Date";
                sheet.Cells["B1"].Value = "Initials";
                sheet.Cells["C1"].Value = "Hours";
                sheet.Cells["D1"].Value = "Note";
            }
            return package.GetAsByteArray();
        }

        public static void Create(string? sheet)
        {
            _worksheet = sheet;
            var reportData = new ReportExcelGenerator().GetReport();
            var reportExcel = new ReportExcelGenerator().Generate(reportData);
            File.WriteAllBytes(@"..\..\..\Classes\Data\EmployeeReports.xlsx", reportExcel);
        }
    }
}