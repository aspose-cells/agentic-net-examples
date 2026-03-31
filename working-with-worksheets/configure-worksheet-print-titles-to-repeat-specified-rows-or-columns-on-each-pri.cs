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
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data
            // First row will act as a header row
            worksheet.Cells["A1"].PutValue("Header 1");
            worksheet.Cells["B1"].PutValue("Header 2");
            worksheet.Cells["C1"].PutValue("Header 3");

            // Fill rows 2‑30 with data; column A will act as a header column
            for (int i = 2; i <= 30; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Row {i}");
                worksheet.Cells[$"B{i}"].PutValue(i * 10);
                worksheet.Cells[$"C{i}"].PutValue(i * 100);
            }

            // Configure page setup to repeat titles
            PageSetup pageSetup = worksheet.PageSetup;

            // Repeat the first row on every printed page
            pageSetup.PrintTitleRows = "$1:$1";

            // Repeat the first column on every printed page
            pageSetup.PrintTitleColumns = "$A:$A";

            // Define the print area (optional, but clarifies what gets printed)
            pageSetup.PrintArea = "A1:C30";

            // Save the workbook
            workbook.Save("PrintTitlesDemo.xlsx");
        }
    }
}