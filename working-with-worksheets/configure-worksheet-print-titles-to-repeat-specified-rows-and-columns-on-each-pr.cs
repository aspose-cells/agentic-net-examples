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

            // Populate some sample data
            // Header rows (to be repeated)
            worksheet.Cells["A1"].PutValue("Header 1");
            worksheet.Cells["B1"].PutValue("Header 2");
            worksheet.Cells["C1"].PutValue("Header 3");

            // Header columns (to be repeated)
            for (int i = 2; i <= 30; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
                worksheet.Cells[$"B{i}"].PutValue(i * 10);
                worksheet.Cells[$"C{i}"].PutValue(i * 100);
            }

            // Access the page setup object
            PageSetup pageSetup = worksheet.PageSetup;

            // Set rows to repeat at the top of each printed page (first row)
            pageSetup.PrintTitleRows = "$1:$1";

            // Set columns to repeat on the left side of each printed page (first column)
            pageSetup.PrintTitleColumns = "$A:$A";

            // Optionally define a print area covering the data
            pageSetup.PrintArea = "A1:C30";

            // Save the workbook
            workbook.Save("PrintTitlesDemo.xlsx");
        }
    }
}