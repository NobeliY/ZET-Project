
using System.Diagnostics;
using OfficeOpenXml;

namespace ZET_Project
{
    public class Test
    {
        public static string? path = @"..\..\..\Classes\Data\EmployeeReports.xlsx";
        public static void AddHours(string tableList)
        {
            var package = new ExcelPackage(path);
            var Sheet = package.Workbook.Worksheets[tableList];
            int lastrow = Sheet.Dimension.End.Row;
            while (Sheet.Cells[lastrow,1].Value == null)
            {
                lastrow--;
            }

            for (int i = 2; i <= lastrow; i++)
            {
                Sheet.Cells[i, 1].Value ??= $"{i+1}.07.2021";
            }
            package.SaveAs("EmployeeReport.xlsx");
            var process = new Process();
            process.StartInfo = new ProcessStartInfo("EmployeeReport.xlsx")
            {
                UseShellExecute = true
            };
            process.Start();
        }

        public static void _Main()
        {
            AddHours("Образец");
        }
    }
}