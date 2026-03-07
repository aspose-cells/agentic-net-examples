using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitlesDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data to demonstrate the effect of print titles
            worksheet.Cells["A1"].PutValue("Header Row");
            for (int i = 2; i <= 30; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Data Row {i - 1}");
                worksheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Access the page setup of the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Set the rows that will repeat at the top of each printed page (e.g., first row)
            pageSetup.PrintTitleRows = "$1:$1";

            // Set the columns that will repeat on the left side of each printed page (e.g., column A)
            pageSetup.PrintTitleColumns = "$A:$A";

            // Save the workbook (lifecycle save)
            workbook.Save("PrintTitlesDemo.xlsx");

            Console.WriteLine("Workbook saved with print titles set.");
        }
    }
}