using System;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterOpNetCore31
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excel;
            Excel.Workbook workbook;
            Excel.Worksheet sheet;

            try
            {
                // Start Excel and get Application object.
                excel = new Excel.Application
                {
                    Visible = true
                };

                // Get a new workbook.
                workbook = excel.Workbooks.Add(Missing.Value);
                sheet = (Excel.Worksheet)workbook.ActiveSheet;

                // Add table headers going cell by cell.
                sheet.Cells[1, 1] = "First Name";
                sheet.Cells[1, 2] = "Last Name";
                sheet.Cells[1, 3] = "Full Name";
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message} Line: {e.Source}");
            }
        }
    }
}
