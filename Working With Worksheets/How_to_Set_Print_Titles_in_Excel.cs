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
            worksheet.Cells["A1"].PutValue("Header Row"); // Row to be repeated
            for (int i = 2; i <= 20; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Title {i - 1}"); // Column to be repeated
                worksheet.Cells[$"B{i}"].PutValue($"Data {i - 1}");
            }

            // Access page setup
            PageSetup pageSetup = worksheet.PageSetup;

            // Set the rows that repeat on each printed page (first row)
            pageSetup.PrintTitleRows = "$1:$1";

            // Set the columns that repeat on each printed page (first column)
            pageSetup.PrintTitleColumns = "$A:$A";

            // Define the print area (optional)
            pageSetup.PrintArea = "A1:B20";

            // Save the workbook
            workbook.Save("PrintTitlesDemo.xlsx");
        }
    }
}