using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitlesDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Add data rows
            for (int i = 2; i <= 30; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i - 1);
                sheet.Cells[$"B{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"C{i}"].PutValue(i * 10);
            }

            // Configure page setup to repeat header row and column on each printed page
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.PrintTitleRows = "$1:$1";   // repeat first row
            pageSetup.PrintTitleColumns = "$A:$A"; // repeat first column
            pageSetup.PrintArea = "A1:C30";       // optional: define the print area

            // Save the workbook
            workbook.Save("PrintTitlesDemo.xlsx");
        }
    }
}