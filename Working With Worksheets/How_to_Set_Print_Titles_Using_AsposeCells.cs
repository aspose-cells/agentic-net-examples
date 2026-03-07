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

            // Add sample data – a header row and several data rows
            worksheet.Cells["A1"].PutValue("Header");
            for (int i = 2; i <= 30; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
                worksheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Set the first row to repeat on every printed page
            pageSetup.PrintTitleRows = "$1:$1";

            // Set the first column to repeat on every printed page
            pageSetup.PrintTitleColumns = "$A:$A";

            // (Optional) Define the print area to limit what gets printed
            pageSetup.PrintArea = "A1:B30";

            // Save the workbook with the print title settings applied
            workbook.Save("PrintTitlesDemo.xlsx");
        }
    }
}