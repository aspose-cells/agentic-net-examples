using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitlesDemo
{
    public class SetPrintTitles
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data to illustrate the effect of print titles
            worksheet.Cells["A1"].PutValue("Header Row");
            for (int i = 2; i <= 30; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Data Row {i - 1}");
                worksheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Access the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Set the rows that will repeat at the top of each printed page (first row)
            pageSetup.PrintTitleRows = "$1:$1";

            // Set the columns that will repeat on the left side of each printed page (first column)
            pageSetup.PrintTitleColumns = "$A:$A";

            // Optionally define a print area to limit what gets printed
            pageSetup.PrintArea = "A1:B30";

            // Save the workbook to a file
            workbook.Save("PrintTitlesDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetPrintTitles.Run();
        }
    }
}